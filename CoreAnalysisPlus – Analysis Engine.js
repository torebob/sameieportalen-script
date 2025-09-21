/**
 * CoreAnalysisPlus – Analysis Engine (v1.9.0, single-file)
 * + Historical metrics, HTML dashboard (Mermaid + D3), rule plugins, JSON API.
 *
 * Public entrypoints:
 *   performComprehensiveAnalysis_(options?)
 *   readRequirementsForGapAnalysis_()
 *   performGapAnalysis_(analysis, existing, newCandidates?, options?)
 *   dedupeCandidates_(candidates, existing, threshold?, options?)
 *   ae_runAndLogAnalysis_(options?)                 // writes a row to Analysis_Log
 *   ae_getLatestMetrics_()                          // compact JSON for CI
 *   ae_openDashboard_()                             // sidebar dashboard
 *   doGet(e)                                        // web app (HTML or JSON API)
 *   validateAnalysisConfig_(), clearAnalysisTokenCache_()
 *
 * Quick start:
 *   ae_runAndLogAnalysis_({ commit: 'abc123' });
 *   ae_openDashboard_(); // sidebar
 *
 * Deploy as Web App (Execute as you; accessible to your domain):
 *   - GET .../exec?format=json&view=latest   // analysis+gap+health
 *   - GET .../exec?format=json&view=history  // all history rows
 *   - GET .../exec?format=json&view=metrics  // compact metrics
 */

/* ---------------------------- Safe dependency bridges ---------------------------- */
function __ae_log_(){try{if(typeof _getLoggerPlus_==='function')return _getLoggerPlus_();}catch(_){}
  return{debug:function(f,m,d){try{console.log('[DEBUG]',f||'',m||'',d||'');}catch(_){}},
         info:function(f,m,d){try{console.log('[INFO]', f||'',m||'',d||'');}catch(_){}},
         warn:function(f,m,d){try{console.warn('[WARN]',f||'',m||'',d||'');}catch(_){}},
         error:function(f,m,d){try{console.error('[ERROR]',f||'',m||'',d||'');}catch(_){}}};}
function __ae_cfgGet_(k,fb){try{if(typeof _cfgGet_==='function')return _cfgGet_(k,fb);}catch(_){}
  try{if(typeof CORE_ANALYSIS_CFG!=='undefined'&&CORE_ANALYSIS_CFG&&Object.prototype.hasOwnProperty.call(CORE_ANALYSIS_CFG,k))return CORE_ANALYSIS_CFG[k];}catch(_){}
  return fb;}
function __ae_numCfg_(k,fb){var v=Number(__ae_cfgGet_(k,fb));return isNaN(v)?Number(fb):v;}
function __ae_safe_(fn,fb){try{return fn();}catch(_){return fb;}}

/* -------------------------------- Config & sanity -------------------------------- */
var __AE_CONST={ LD_SHEETS:__ae_numCfg_('LARGE_DATA_SHEETS',12),
  LD_MAXCOLS:__ae_numCfg_('LARGE_DATA_MAXCOLS',60),
  LD_TOTALROWS:__ae_numCfg_('LARGE_DATA_TOTALROWS',50000),
  VERSION:__ae_cfgGet_('VERSION','1.9.0'),
  TOKEN_MIN:__ae_numCfg_('TOKEN_MIN_LEN',2),
  JACCARD_DEF:__ae_numCfg_('DEFAULT_JACCARD_THRESHOLD',0.78),
  DEDUPE_BATCH:__ae_numCfg_('DEDUPE_BATCH_SIZE',400),
  LOG_SHEET:'Analysis_Log',
  DASH_TITLE:'Analyse – CoreAnalysisPlus' };
function validateAnalysisConfig_(){
  var issues=[];
  function rng(n,v,a,b){if(typeof v!=='number'||isNaN(v)||v<a||v>b)issues.push(n+' out of range: '+v);}
  function pos(n,v){if(!Number.isFinite(v)||v<=0)issues.push(n+' must be > 0');}
  rng('DEFAULT_JACCARD_THRESHOLD',__AE_CONST.JACCARD_DEF,0,1);
  pos('TOKEN_MIN_LEN',__AE_CONST.TOKEN_MIN);
  pos('DEDUPE_BATCH_SIZE',__AE_CONST.DEDUPE_BATCH);
  if(issues.length){__ae_log_().warn('validateAnalysisConfig_', 'Non-fatal config issues', {issues:issues});}
  return issues;
}

/* -------------------------------- Token & Jaccard -------------------------------- */
var __ae_tokenCache=Object.create(null);
function clearAnalysisTokenCache_(){__ae_tokenCache=Object.create(null);}
function __ae_splitCamel_(s){return String(s||'').replace(/([a-z0-9])([A-Z])/g,'$1 $2').replace(/[_\-\.]+/g,' ');}
function __ae_tokens_(s){var key='t:'+s;if(__ae_tokenCache[key])return __ae_tokenCache[key];
  var t=__ae_splitCamel_(s).toLowerCase().split(/[^a-z0-9æøå]+/).filter(function(x){return x&&x.length>=__AE_CONST.TOKEN_MIN;});
  __ae_tokenCache[key]=t;return t;}
function __ae_jaccard_(a,b,minTh){var A=__ae_tokens_(a),B=__ae_tokens_(b);
  if(A.length===0&&B.length===0)return 1; if(A.length===0||B.length===0)return 0;
  var minLen=Math.min(A.length,B.length),maxLen=Math.max(A.length,B.length);
  var upper=minLen/maxLen; if(typeof minTh==='number'&&upper<minTh)return 0;
  var setA=Object.create(null); for(var i=0;i<A.length;i++)setA[A[i]]=1;
  var inter=0, union=Object.create(null); for(var k in setA)union[k]=1;
  for(var j=0;j<B.length;j++){var tok=B[j]; if(!union[tok])union[tok]=1; if(setA[tok])inter++;}
  var u=0; for(var x in union)u++; return u? inter/u : 0;}
function __ae_jaccardThreshold_(p){if(typeof p==='number'&&!isNaN(p))return Math.min(1,Math.max(0,p));
  try{if(typeof CONFIG_PLUS!=='undefined'&&CONFIG_PLUS){
    if(typeof CONFIG_PLUS.DEFAULT_JACCARD_THRESHOLD==='number')return CONFIG_PLUS.DEFAULT_JACCARD_THRESHOLD;
    if(typeof CONFIG_PLUS.DEDUPE_JACCARD==='number')return CONFIG_PLUS.DEDUPE_JACCARD;}}catch(_){}
  return __AE_CONST.JACCARD_DEF;}
function __ae_nowISO_(){return new Date().toISOString();}
function __ae_progress_(cb,p){if(typeof cb==='function'){try{cb(p);}catch(_){}}}

/* ---------------------------- Inventory merge (helpers) --------------------------- */
function __ae_mergeFunctions_(triggers,menus){
  var seen=Object.create(null), out=[];
  (triggers||[]).forEach(function(t){var n=String(t.handler||'').trim(); if(!n||seen[n])return; seen[n]=1; out.push({name:n,source:'trigger',eventType:String(t.eventType||'')});});
  (menus||[]).forEach(function(m){var n=String(m.functionName||'').trim(); if(!n||seen[n])return; seen[n]=1; out.push({name:n,source:'menu',title:String(m.title||'')});});
  return out;
}

/* ------------------------------------ Orchestrator ------------------------------------ */
function performComprehensiveAnalysis_(options){
  validateAnalysisConfig_();
  var started=Date.now(), progressCb=options&&options.progressCb;
  var meta=__ae_safe_(_collectMetadata_,{}),
      triggers=__ae_safe_(_collectTriggers_,[]),
      menus=__ae_safe_(_collectMenuFunctions_,[]),
      model={sheets:[],headerDuplicates:[]};

  try{
    if(typeof _collectDataModel_==='function' && _collectDataModel_.length>=1){
      model=_collectDataModel_(function(p){__ae_progress_(progressCb,{phase:'dataModel', sheetName:p&&p.sheetName, current:p&&p.current, total:p&&p.total, percentage:p&&p.percentage});});
    }else{
      model=__ae_safe_(_collectDataModel_,{sheets:[],headerDuplicates:[]});
    }
  }catch(_){}

  var functions=__ae_mergeFunctions_(triggers,menus);
  var arr=model.sheets||[], sheets=arr.length, rows=0, maxCols=0, cells=0;
  for(var i=0;i<arr.length;i++){var r=Number(arr[i].rows||0), c=Number(arr[i].columns||0); rows+=r; if(c>maxCols)maxCols=c; cells+=r*c;}
  var large=(sheets>=__AE_CONST.LD_SHEETS)||(maxCols>=__AE_CONST.LD_MAXCOLS)||(rows>=__AE_CONST.LD_TOTALROWS);
  var tokenKeys=0; try{tokenKeys=Object.keys(__ae_tokenCache).length;}catch(_){}
  return { version:__AE_CONST.VERSION, timestamp:__ae_nowISO_(), metadata:meta,
    triggers:{count:(triggers||[]).length,details:triggers||[]},
    menus:{fromSheets:menus||[]},
    functions:{global:functions,private:[]},
    sheets:{count:sheets, sheets:arr, headerDuplicates:model.headerDuplicates||[]},
    performanceMetrics:{sheetsScanned:sheets,totalRows:rows,maxCols:maxCols,totalCells:cells,scanDurationMs:Date.now()-started,largeDataset:large,tokenCacheKeys:tokenKeys}};
}

/* ------------------------------ Requirements reader ------------------------------ */
function readRequirementsForGapAnalysis_(){
  var names=__ae_cfgGet_('NAMES',{kravSheet:['Krav','Requirements']});
  var sh=(typeof _getSheetByAnyName_==='function')?_getSheetByAnyName_(names.kravSheet):null;
  if(!sh) return [];
  var vals=sh.getDataRange().getValues(); if(!vals||vals.length<2) return [];
  var headers=vals[0].map(function(h){return String(h||'').trim().toLowerCase();});
  var KH=__ae_cfgGet_('HEADERS',{krav:{id:['krav id','kravid','id','krav-id'],text:['krav','beskrivelse','tekst','requirement','description','text'],priority:['prioritet','prio','priority'],progress:['fremdrift %','fremdrift%','fremdrift','progress','%'],chapter:['kapittel','kap','chapter']}}).krav;
  function idxAny(alts){for(var i=0;i<headers.length;i++){for(var j=0;j<alts.length;j++){if(headers[i]===String(alts[j]).toLowerCase())return i;}} return -1;}
  var idI=idxAny(KH.id), tI=idxAny(KH.text), pI=idxAny(KH.priority), prI=idxAny(KH.progress), cI=idxAny(KH.chapter);
  var out=[]; for(var r=1;r<vals.length;r++){var row=vals[r]||[];
    out.push({id:(idI>=0?String(row[idI]||''):''), text:(tI>=0?String(row[tI]||''):''), priority:(pI>=0?String(row[pI]||''):''), progressPct:Number((prI>=0?row[prI]:0)||0), chapter:(cI>=0?String(row[cI]||''):'')});}
  return out;
}

/* ------------------------------------- Dedupe ------------------------------------- */
function dedupeCandidates_(candidates,existing,threshold,options){
  var log=__ae_log_(); if(!Array.isArray(candidates)) {log.warn('dedupeCandidates_','candidates must be array'); return [];}
  existing=Array.isArray(existing)?existing:[]; options=options||{};
  var batchSize=Math.max(50, Number(options.batchSize||__AE_CONST.DEDUPE_BATCH)||__AE_CONST.DEDUPE_BATCH);
  var progressCb=options.progressCb, th=__ae_jaccardThreshold_(threshold);
  var existTexts=existing.map(function(e){return String(e&&e.text||'');}).filter(Boolean);
  var out=[], seen=[], processed=0, total=candidates.length;
  for(var start=0; start<total; start+=batchSize){
    var end=Math.min(start+batchSize,total);
    for(var i=start;i<end;i++){
      var t=String(candidates[i].text||''); if(!t){processed++;continue;}
      var dupExist=false; for(var ex=0;ex<existTexts.length;ex++){ if(__ae_jaccard_(t,existTexts[ex],th)>=th){dupExist=true;break;}}
      if(dupExist){processed++;continue;}
      var dupNew=false; for(var s=0;s<seen.length;s++){ if(__ae_jaccard_(t,seen[s],th)>=th){dupNew=true;break;}}
      if(dupNew){processed++;continue;}
      seen.push(t); out.push(candidates[i]); processed++;
    }
    __ae_progress_(progressCb,{phase:'dedupe',processed:processed,total:total,percent:Math.round((processed/Math.max(1,total))*100)});
    try{if(start>0&&(start/batchSize)%4===0)Utilities.sleep(5);}catch(_){}
  }
  return out;
}

/* --------------------------------- Gap & Coverage -------------------------------- */
function __ae_semanticCovered_(fnName,lowerTexts,threshold){
  var f=String(fnName||''); if(!f)return false; var th=__ae_jaccardThreshold_(threshold); var lower=f.toLowerCase();
  for(var i=0;i<lowerTexts.length;i++){ if(lowerTexts[i].indexOf(lower)>=0) return true; }
  for(var j=0;j<lowerTexts.length;j++){ if(__ae_jaccard_(f,lowerTexts[j],th)>=th) return true; }
  return false;
}
function performGapAnalysis_(analysis, existing, newCandidates, options){
  options=options||{}; var A=analysis||{}, reqs=Array.isArray(existing)?existing:[], cands=Array.isArray(newCandidates)?newCandidates:[];
  var impl0=[], partial=[], done=[]; for(var i=0;i<reqs.length;i++){var p=Number(reqs[i].progressPct||0); if(p<=0)impl0.push(reqs[i]); else if(p>=100)done.push(reqs[i]); else partial.push(reqs[i]);}
  var publicFns=(A.functions&&A.functions.global?A.functions.global:[]).map(function(f){return String(f.name||'');}).filter(Boolean);
  var kravL=reqs.map(function(r){return String(r.text||'').toLowerCase();});
  var candL=cands.map(function(c){return String(c.text||'').toLowerCase();});
  var th=(typeof options.coverageThreshold==='number')?options.coverageThreshold:__AE_CONST.JACCARD_DEF;
  var undocumented=[]; for(var f=0;f<publicFns.length;f++){var fn=publicFns[f]; var cov=__ae_semanticCovered_(fn,kravL,th)||__ae_semanticCovered_(fn,candL,th); if(!cov)undocumented.push({function:fn});}
  var total=Math.max(1,reqs.length);
  var coverage={ total: total, implemented: done.length, partial: partial.length, missing: impl0.length,
    implementedPct: Math.round((done.length/total)*100),
    coveredOrPartialPct: Math.round(((done.length+partial.length)/total)*100) };
  function txtScore(t){var s=String(t||'').trim(); if(!s)return 0; var len=s.length; var lb=(len>=50&&len<=240)?1:0.5; var punct=/[.!?]$/.test(s)?1:0.8; return Math.round((lb*0.6+punct*0.4)*100)/100;}
  var recs=[]; if(coverage.missing>0)recs.push({rank:1,message:'Fiks MÅ-krav med 0% fremdrift først.',weight:0.9});
  if(undocumented.length>0)recs.push({rank:2,message:'Dokumenter funksjoner uten dekning (semantikk).',weight:0.85});
  if(partial.length>0)recs.push({rank:3,message:'Fullfør delvis implementerte krav (raske gevinster).',weight:0.7});
  recs.sort(function(a,b){return b.weight-a.weight;});
  var candQuality=cands.slice(0,50).map(function(c){return{text:c.text,quality:txtScore(c.text)};});
  var tokenKeys=0; try{tokenKeys=Object.keys(__ae_tokenCache).length;}catch(_){}
  return{coverage:coverage,undocumentedFunctions:undocumented,unimplementedRequirements:impl0,partialRequirements:partial,implementedRequirements:done,candidateQualitySample:candQuality,recommendations:recs,telemetry:{tokenCacheKeys:tokenKeys}};
}

/* ------------------------------- Rule Plugin System ------------------------------- */
var __ae_rules = [];
__ae_rules.push(function rule_duplicateFunctions(ctx){
  var fns=(ctx.analysis.functions&&ctx.analysis.functions.global)||[];
  var th=0.95, findings=[];
  for(var i=0;i<fns.length;i++){
    for(var j=i+1;j<fns.length;j++){
      var a=fns[i].name, b=fns[j].name;
      if(__ae_jaccard_(a,b,th)>=th){ findings.push({type:'duplicate-function', a:a, b:b, similarity:'>=0.95'}); }
    }
  }
  return findings;
});
__ae_rules.push(function rule_deadCode(ctx){
  var fns=(ctx.analysis.functions&&ctx.analysis.functions.global)||[];
  var kravL=(ctx.requirements||[]).map(function(r){return String(r.text||'').toLowerCase();});
  var menuL=(ctx.analysis.menus&&ctx.analysis.menus.fromSheets||[]).map(function(m){return String(m.title||'').toLowerCase();});
  var out=[];
  for(var i=0;i<fns.length;i++){
    var n=String(fns[i].name||''); if(!n) continue;
    var low=n.toLowerCase();
    var hinted=kravL.some(function(t){return t.indexOf(low)>=0;}) || menuL.some(function(t){return t.indexOf(low)>=0;});
    if(!hinted) out.push({type:'dead-code-heuristic', function:n, note:'Ikke referert i krav/meny'});
  }
  return out;
});
function ae_runRules_(analysis, requirements, gap){
  var ctx={analysis:analysis, requirements:requirements, gap:gap};
  var all=[]; for(var i=0;i<__ae_rules.length;i++){ try{var f=__ae_rules[i](ctx)||[]; all=all.concat(f);}catch(e){__ae_log_().warn('ae_runRules_', 'rule failed', {i:i, err:e&&e.message});}}
  return all;
}

/* -------------------------------- Code Health Score ------------------------------- */
function ae_codeHealthSummary_(analysis, gap){
  var perf=analysis.performanceMetrics||{};
  var undocumented=(gap.undocumentedFunctions||[]).length;
  var implPct=gap.coverage && gap.coverage.implementedPct || 0;
  var score=Math.max(0, Math.min(100,
    (implPct*0.6) +
    ((100 - Math.min(undocumented*5,60))*0.3) +
    ((perf.largeDataset? 5 : 10))*0.1
  ));
  var grade = score>=90?'A' : score>=80?'B' : score>=70?'C' : score>=60?'D' : 'E';
  return { score: Math.round(score), grade: grade, implementedPct: implPct, undocumentedCount: undocumented };
}

/* ------------------------------- Historical persistence ---------------------------- */
function __ae_getOrCreateLog_(){
  var ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName(__AE_CONST.LOG_SHEET);
  if(!sh) sh=ss.insertSheet(__AE_CONST.LOG_SHEET);
  if(sh.getLastRow()===0){
    sh.appendRow(['Timestamp','Version','Commit','Sheets','Rows','MaxCols','Cells','ScanMs','Undocumented','ImplementedPct','Grade','Score']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function ae_runAndLogAnalysis_(options){
  options=options||{};
  var analysis=performComprehensiveAnalysis_({progressCb:options.progressCb});
  var reqs=readRequirementsForGapAnalysis_();
  var gap=performGapAnalysis_(analysis, reqs, [], {});
  var health=ae_codeHealthSummary_(analysis, gap);
  var sh=__ae_getOrCreateLog_();
  var perf=analysis.performanceMetrics||{};
  var row=[ new Date(), analysis.version, String(options.commit||''),
    perf.sheetsScanned||0, perf.totalRows||0, perf.maxCols||0, perf.totalCells||0, perf.scanDurationMs||0,
    (gap.undocumentedFunctions||[]).length, health.implementedPct, health.grade, health.score ];
  sh.appendRow(row);
  return {analysis:analysis, gap:gap, health:health, logRow: sh.getLastRow()};
}
function ae_getLatestMetrics_(){
  var sh=__ae_getOrCreateLog_(); var r=sh.getLastRow(); if(r<2) return null;
  var vals=sh.getRange(r,1,1,sh.getLastColumn()).getValues()[0];
  return { timestamp: vals[0], version: vals[1], commit: vals[2],
    sheets: vals[3], rows: vals[4], maxCols: vals[5], cells: vals[6], scanMs: vals[7],
    undocumented: vals[8], implementedPct: vals[9], grade: vals[10], score: vals[11] };
}

/* --------------------------- Dependency graph (Mermaid) --------------------------- */
/**
 * Build a simple dependency graph: Functions, Sheets, Triggers, Menus.
 * Returns Mermaid flowchart string (graph LR).
 */
function ae_buildMermaid_(analysis){
  var A=analysis||{};
  var lines=['graph LR'];
  function safeId(s){return (String(s||'').replace(/[^a-zA-Z0-9_]/g,'_')||'X');}

  // Nodes
  var sheets=(A.sheets&&A.sheets.sheets)||[];
  var funcs=(A.functions&&A.functions.global)||[];
  var trigs=(A.triggers&&A.triggers.details)||[];
  var menus=(A.menus&&A.menus.fromSheets)||[];

  // Declare nodes with types
  for(var i=0;i<sheets.length;i++){
    var sn=String(sheets[i].name||''); if(!sn)continue;
    lines.push(safeId('sheet_'+sn)+'["Sheet: '+sn+'"]:::sheet');
  }
  for(var j=0;j<funcs.length;j++){
    var fn=String(funcs[j].name||''); if(!fn)continue;
    lines.push(safeId('fn_'+fn)+'(("Fn: '+fn+'")):::fn');
  }
  for(var k=0;k<trigs.length;k++){
    var et=String(trigs[k].eventType||'UNKNOWN'), hf=String(trigs[k].handler||'');
    lines.push(safeId('trig_'+hf)+'["Trig: '+et+'"]:::trig');
  }
  for(var m=0;m<menus.length;m++){
    var mt=String(m.title||m.functionName||'');
    lines.push(safeId('menu_'+mt)+'["Menu: '+mt+'"]:::menu');
  }

  // Edges: triggers -> function
  for(var t=0;t<trigs.length;t++){
    var h=String(trigs[t].handler||''); if(!h) continue;
    lines.push(safeId('trig_'+h)+' --> '+safeId('fn_'+h));
  }
  // Edges: menu -> function
  for(var mm=0;mm<menus.length;mm++){
    var fnn=String(menus[mm].functionName||''); var ttl=String(menus[mm].title||menus[mm].functionName||'');
    if(fnn) lines.push(safeId('menu_'+ttl)+' --> '+safeId('fn_'+fnn));
  }
  // Heuristic: function -> sheet (if sheet name appears in header preview or vice versa)
  for(var f=0;f<funcs.length;f++){
    var fname=String(funcs[f].name||''); if(!fname)continue;
    var fl=fname.toLowerCase();
    for(var s=0;s<sheets.length;s++){
      var sname=String(sheets[s].name||''); var hp=String(sheets[s].headerPreview||'').toLowerCase();
      if(hp.indexOf(fl)>=0 || sname.toLowerCase().indexOf(fl)>=0){
        lines.push(safeId('fn_'+fname)+' --> '+safeId('sheet_'+sname));
      }
    }
  }

  // Styles
  lines.push('classDef sheet fill:#ECFCCB,stroke:#84CC16,color:#1a1a1a;');
  lines.push('classDef fn fill:#E0F2FE,stroke:#38BDF8,color:#1a1a1a;');
  lines.push('classDef trig fill:#FFE4E6,stroke:#FB7185,color:#1a1a1a;');
  lines.push('classDef menu fill:#F3E8FF,stroke:#A78BFA,color:#1a1a1a;');
  return lines.join('\n');
}

/* ---------------------------------- Dashboard UI ---------------------------------- */
function ae_openDashboard_(){
  var html=HtmlService.createHtmlOutput(__ae_dashboardHTML_()).setTitle(__AE_CONST.DASH_TITLE).setWidth(520);
  SpreadsheetApp.getUi().showSidebar(html);
}
function doGet(e){
  var p=(e&&e.parameter)||{};
  if(String(p.format||'')==='json'){
    var view=String(p.view||'latest');
    var payload=null;
    if(view==='history'){ payload=__ae_getHistory_(); }
    else if(view==='metrics'){ payload=ae_getLatestMetrics_(); }
    else { // latest
      var latest = ae_runPreview_();
      payload = latest ? { analysis:latest.analysis, gap:latest.gap, health:latest.health } : null;
    }
    return ContentService.createTextOutput(JSON.stringify(payload||null)).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutput(__ae_dashboardHTML_()).setTitle(__AE_CONST.DASH_TITLE);
}
function __ae_dashboardHTML_(){
  var title=__AE_CONST.DASH_TITLE;
  return `
<!DOCTYPE html><html><head><meta charset="utf-8" />
<title>${title}</title>
<style>
  body{font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;margin:12px;}
  h1{font-size:18px;margin:0 0 8px;}
  .kpi{display:flex;gap:8px;flex-wrap:wrap;margin:8px 0;}
  .card{border:1px solid #e5e7eb;border-radius:10px;padding:10px;flex:1 1 140px;min-width:140px}
  .muted{color:#6b7280}
  .row{display:flex;gap:8px;align-items:center;margin:8px 0}
  .pill{padding:2px 8px;border-radius:999px;background:#f3f4f6;font-size:11px}
  .tabs{display:flex;gap:6px;margin:8px 0}
  .tab{padding:6px 10px;border:1px solid #e5e7eb;border-radius:8px;background:#fff;cursor:pointer}
  .tab.active{background:#eef2ff;border-color:#c7d2fe}
  table{border-collapse:collapse;width:100%;font-size:12px}
  th,td{border:1px solid #eee;padding:6px;text-align:left}
  th{background:#fafafa;cursor:pointer}
  #hist{height:180px;border:1px solid #eee;border-radius:8px;padding:4px}
  #graph{border:1px solid #eee;border-radius:8px;padding:8px;max-height:420px;overflow:auto}
</style>
</head><body>
  <h1>${title}</h1>
  <div class="row">
    <button class="tab active" id="tDash">Dashboard</button>
    <button class="tab" id="tUndoc">Undocumented</button>
    <button class="tab" id="tGraph">Dependency Graph</button>
    <span id="status" class="muted" style="margin-left:auto"></span>
  </div>

  <div id="vDash">
    <div class="kpi">
      <div class="card"><div class="muted">Karakter</div><div id="grade" style="font-size:24px">–</div></div>
      <div class="card"><div class="muted">Score</div><div id="score" style="font-size:24px">–</div></div>
      <div class="card"><div class="muted">Implementert</div><div id="impl" style="font-size:24px">–</div></div>
      <div class="card"><div class="muted">Udekket</div><div id="undoc" style="font-size:24px">–</div></div>
    </div>

    <div class="row"><div class="muted">Historikk</div><span class="pill" id="histCount">–</span>
      <button class="tab" style="margin-left:auto" onclick="runNow()">Kjør analyse nå</button>
    </div>
    <div id="hist" title="Implementert % over tid">Laster historikk…</div>

    <h3>Foreslåtte tiltak</h3>
    <ul id="recs"></ul>
  </div>

  <div id="vUndoc" style="display:none">
    <h3>Udekkede funksjoner (semantisk)</h3>
    <table id="tblUndoc"><thead><tr><th onclick="sortUndoc(0)">Funksjon</th></tr></thead><tbody></tbody></table>
  </div>

  <div id="vGraph" style="display:none">
    <h3>Avhengighetsgraf</h3>
    <div id="graph">Laster graf…</div>
  </div>

<script>
  function setStatus(t){document.getElementById('status').textContent=t||'';}
  function $(id){return document.getElementById(id);}
  function show(tab){
    $('tDash').classList.toggle('active',tab==='dash'); $('vDash').style.display=(tab==='dash'?'block':'none');
    $('tUndoc').classList.toggle('active',tab==='undoc'); $('vUndoc').style.display=(tab==='undoc'?'block':'none');
    $('tGraph').classList.toggle('active',tab==='graph'); $('vGraph').style.display=(tab==='graph'?'block':'none');
  }
  $('tDash').onclick=()=>show('dash');
  $('tUndoc').onclick=()=>show('undoc');
  $('tGraph').onclick=()=>show('graph');

  function renderHistoryD3(hist){
    const root=$('hist'); root.innerHTML='';
    if(!hist||hist.length===0){root.textContent='Ingen data ennå'; return;}
    // Try load D3, fallback to simple bars if blocked.
    const scr=document.createElement('script');
    scr.src='https://cdn.jsdelivr.net/npm/d3@7';
    scr.onload=()=>drawD3(hist);
    scr.onerror=()=>drawFallback(hist);
    root.appendChild(scr);

    function drawD3(data){
      const vals=data.slice(-40).map(r=>Number(r[9]||0)); // ImplementedPct
      const w=root.clientWidth-8, h=root.clientHeight-8;
      const svg=d3.select(root).append('svg').attr('width',w).attr('height',h);
      const x=d3.scaleBand().domain(d3.range(vals.length)).range([0,w]).padding(0.1);
      const y=d3.scaleLinear().domain([0,100]).range([h,0]);
      svg.selectAll('rect').data(vals).enter().append('rect')
        .attr('x',(d,i)=>x(i)).attr('y',d=>y(d))
        .attr('width',x.bandwidth()).attr('height',d=>h-y(d)).attr('fill','#60a5fa');
    }
    function drawFallback(data){
      const wrap=document.createElement('div'); wrap.style.display='flex'; wrap.style.alignItems='flex-end'; wrap.style.gap='4px'; wrap.style.height='100%';
      const vals=data.slice(-40).map(r=>Number(r[9]||0));
      vals.forEach(v=>{const b=document.createElement('div'); b.style.width='10px'; b.style.height=Math.max(6,(v/100)*160)+'px'; b.style.background='#60a5fa'; wrap.appendChild(b);});
      root.appendChild(wrap);
    }
  }

  function renderGraphMermaid(mmd){
    const root=$('graph'); root.innerHTML='';
    const scr=document.createElement('script');
    scr.src='https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js';
    scr.onload=()=>{
      try { mermaid.initialize({ startOnLoad:false, theme:'default', securityLevel:'loose' }); } catch(e){}
      const div=document.createElement('div'); div.className='mermaid'; div.textContent=mmd;
      root.appendChild(div);
      try { mermaid.init(undefined, div); } catch(e){ root.textContent='Kunne ikke render graf.'; }
    };
    scr.onerror=()=>{
      root.textContent='Mermaid kunne ikke lastes (nettverk/CSP).';
      const pre=document.createElement('pre'); pre.textContent=mmd; root.appendChild(pre);
    };
    root.appendChild(scr);
  }

  function loadLatest(){
    setStatus('Laster…');
    google.script.run.withSuccessHandler(function(payload){
      setStatus('');
      if(!payload){$('grade').textContent='–'; return;}
      $('grade').textContent=payload.health.grade;
      $('score').textContent=payload.health.score;
      $('impl').textContent=payload.health.implementedPct+'%';
      $('undoc').textContent=payload.gap.undocumentedFunctions.length;
      const recs=$('recs'); recs.innerHTML='';
      (payload.gap.recommendations||[]).forEach(r=>{const li=document.createElement('li'); li.textContent=r.message; recs.appendChild(li);});
      // History & chart
      google.script.run.withSuccessHandler(function(hist){ $('histCount').textContent=String(hist.length); renderHistoryD3(hist); }).__ae_getHistory__();
      // Undocumented table
      const tb=$('tblUndoc').querySelector('tbody'); tb.innerHTML='';
      (payload.gap.undocumentedFunctions||[]).slice(0,500).forEach(u=>{const tr=document.createElement('tr'); const td=document.createElement('td'); td.textContent=u.function; tr.appendChild(td); tb.appendChild(tr);});
      // Graph
      google.script.run.withSuccessHandler(function(mmd){ renderGraphMermaid(mmd); }).__ae_mermaidForLatest__();
    }).__ae_dashboardData__();
  }
  function runNow(){
    setStatus('Kjører analyse…');
    google.script.run.withSuccessHandler(function(){ setStatus('Ferdig'); loadLatest(); })
      .ae_runAndLogAnalysis_({});
  }
  // bootstrap
  loadLatest();
</script>
</body></html>`;
}
// Data providers for dashboard
function __ae_dashboardData__(){
  var latest = ae_runPreview_();
  return latest ? { analysis:latest.analysis, gap:latest.gap, health:latest.health } : null;
}
function ae_runPreview_(){
  try{
    var analysis=performComprehensiveAnalysis_();
    var reqs=readRequirementsForGapAnalysis_();
    var gap=performGapAnalysis_(analysis, reqs, [], {});
    var health=ae_codeHealthSummary_(analysis, gap);
    return {analysis:analysis,gap:gap,health:health};
  }catch(e){ __ae_log_().warn('ae_runPreview_', 'preview failed', {err:e&&e.message}); return null;}
}
function __ae_getHistory_(){
  var ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName(__AE_CONST.LOG_SHEET); if(!sh) return [];
  if(sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-2+1, sh.getLastColumn()).getValues();
}
function __ae_mermaidForLatest__(){
  var latest=ae_runPreview_(); if(!latest) return 'graph LR\nEmpty["Ingen data"]';
  return ae_buildMermaid_(latest.analysis);
}

/* ---------------------------------- CI helpers ------------------------------------ */
function ae_ciPrintLatest_(){
  var m=ae_getLatestMetrics_(); Logger.log(JSON.stringify(m));
}

/* ---------------------------------- Smoke runner ---------------------------------- */
function runCoreAnalysis_Smoke(){
  var out=ae_runAndLogAnalysis_({});
  __ae_log_().info('runCoreAnalysis_Smoke','ok',{row:out.logRow,health:out.health});
}
