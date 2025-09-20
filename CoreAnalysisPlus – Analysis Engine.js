/**
 * CoreAnalysisPlus – Analysis Engine (v1.8.0, single-file)
 * + Historical metrics, HTML dashboard, rule plugin system, CI helpers.
 *
 * Public entrypoints you’ll use:
 *   performComprehensiveAnalysis_(options?)
 *   readRequirementsForGapAnalysis_()
 *   performGapAnalysis_(analysis, existing, newCandidates?, options?)
 *   dedupeCandidates_(candidates, existing, threshold?, options?)
 *   ae_runAndLogAnalysis_(options?)                 // ← write historical row
 *   ae_getLatestMetrics_()                          // ← compact JSON for CI
 *   ae_openDashboard_()                             // ← open sidebar dashboard
 *   doGet(e)                                        // ← webapp dashboard
 *   validateAnalysisConfig_(), clearAnalysisTokenCache_()
 *
 * Quick start:
 *   ae_runAndLogAnalysis_({ commit: 'abc123' });  // logs a row in Analysis_Log
 *   ae_openDashboard_();                          // sidebar UI
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
  VERSION:__ae_cfgGet_('VERSION','1.8.0'),
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
  var log=__ae_log_(), started=Date.now(), progressCb=options&&options.progressCb;
  var meta=__ae_safe_(_collectMetadata_,{}),
      triggers=__ae_safe_(_collectTriggers_,[]),
      menus=__ae_safe_(_collectMenuFunctions_,[]),
      model={sheets:[],headerDuplicates:[]};

  // Prefer progress-aware DataCollector if available
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
  var res={ version:__AE_CONST.VERSION, timestamp:__ae_nowISO_(), metadata:meta,
    triggers:{count:(triggers||[]).length,details:triggers||[]},
    menus:{fromSheets:menus||[]},
    functions:{global:functions,private:[]},
    sheets:{count:sheets, sheets:arr, headerDuplicates:model.headerDuplicates||[]},
    performanceMetrics:{sheetsScanned:sheets,totalRows:rows,maxCols:maxCols,totalCells:cells,scanDurationMs:Date.now()-started,largeDataset:large,tokenCacheKeys:tokenKeys}};
  __ae_log_().info('performComprehensiveAnalysis_','done',{perf:res.performanceMetrics,fn:functions.length});
  return res;
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
// Add new analysis rules without touching core (accepts {analysis,gap}, returns findings[])
var __ae_rules = [];

// Rule #1: Near-duplicate functions (semantic)
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

// Rule #2: Dead-code heuristic (not referenced by triggers/menus/requirements)
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
  // Simple weighted health score 0..100
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

/**
 * Run analysis, compute gap + score, append to Analysis_Log.
 * @param {{commit?:string, progressCb?:function}} [options]
 * @return {{analysis:Object,gap:Object,health:Object,logRow:number}}
 */
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

/**
 * Compact JSON for CI/CD (emit via logger or WebApp fetch)
 */
function ae_getLatestMetrics_(){
  var sh=__ae_getOrCreateLog_(); var r=sh.getLastRow(); if(r<2) return null;
  var vals=sh.getRange(r,1,1,sh.getLastColumn()).getValues()[0];
  return { timestamp: vals[0], version: vals[1], commit: vals[2],
    sheets: vals[3], rows: vals[4], maxCols: vals[5], cells: vals[6], scanMs: vals[7],
    undocumented: vals[8], implementedPct: vals[9], grade: vals[10], score: vals[11] };
}

/* ---------------------------------- Dashboard UI ---------------------------------- */
// Open as sidebar inside Sheets
function ae_openDashboard_(){
  var html=HtmlService.createHtmlOutput(__ae_dashboardHTML_()).setTitle(__AE_CONST.DASH_TITLE).setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}
// Web app entry (Deploy → Web app). Also serves same dashboard.
function doGet(e){ return HtmlService.createHtmlOutput(__ae_dashboardHTML_()).setTitle(__AE_CONST.DASH_TITLE); }

function __ae_dashboardHTML_(){
  var title=__AE_CONST.DASH_TITLE;
  return `
<!DOCTYPE html><html><head><meta charset="utf-8" />
<title>${title}</title>
<style>
  body{font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;margin:12px;}
  h1{font-size:18px;margin:0 0 8px;}
  .kpi{display:flex;gap:8px;flex-wrap:wrap;margin:8px 0;}
  .card{border:1px solid #e5e7eb;border-radius:10px;padding:10px;flex:1 1 120px}
  .muted{color:#6b7280}
  table{border-collapse:collapse;width:100%;font-size:12px}
  th,td{border:1px solid #eee;padding:6px;text-align:left}
  th{background:#fafafa;cursor:pointer}
  .btn{padding:6px 10px;border:1px solid #e5e7eb;border-radius:8px;background:#fff;cursor:pointer}
  .btn:active{transform:translateY(1px)}
  .row{display:flex;gap:8px;align-items:center;margin:8px 0}
  .pill{padding:2px 8px;border-radius:999px;background:#f3f4f6;font-size:11px}
  #chart{height:160px;margin:8px 0;border:1px solid #eee;border-radius:8px;display:flex;align-items:flex-end;padding:8px;gap:4px}
  .bar{background:#93c5fd;width:16px}
</style>
</head><body>
  <h1>${title}</h1>
  <div class="row">
    <button class="btn" onclick="runNow()">Kjør analyse nå</button>
    <span id="status" class="muted"></span>
  </div>
  <div class="kpi">
    <div class="card"><div class="muted">Karakter</div><div id="grade" style="font-size:24px">–</div></div>
    <div class="card"><div class="muted">Score</div><div id="score" style="font-size:24px">–</div></div>
    <div class="card"><div class="muted">Implementert</div><div id="impl" style="font-size:24px">–</div></div>
    <div class="card"><div class="muted">Udekket</div><div id="undoc" style="font-size:24px">–</div></div>
  </div>

  <div class="row"><div class="muted">Historikk</div><span class="pill" id="histCount">–</span></div>
  <div id="chart" title="Implementert % over tid"></div>

  <h3>Foreslåtte tiltak</h3>
  <ul id="recs"></ul>

  <h3>Udekkede funksjoner (semantisk)</h3>
  <table id="tblUndoc"><thead><tr><th onclick="sortUndoc(0)">Funksjon</th></tr></thead><tbody></tbody></table>

<script>
  function setStatus(t){document.getElementById('status').textContent=t||'';}
  function $(id){return document.getElementById(id);}
  function renderChart(hist){
    const c=$('chart'); c.innerHTML='';
    if(!hist || hist.length===0){ c.textContent='Ingen data ennå'; return; }
    const max=100;
    hist.slice(-20).forEach(r=>{
      const v = Number(r[9]||0); // ImplementedPct col
      const bar=document.createElement('div'); bar.className='bar';
      bar.style.height=(Math.max(5,(v/max)*100))+'%';
      bar.title = new Date(r[0]).toLocaleString()+': '+v+'%';
      c.appendChild(bar);
    });
    $('histCount').textContent = String(hist.length);
  }
  function loadLatest(){
    setStatus('Laster...');
    google.script.run.withSuccessHandler(function(payload){
      setStatus('');
      if(!payload){return;}
      $('grade').textContent=payload.health.grade;
      $('score').textContent=payload.health.score;
      $('impl').textContent=payload.health.implementedPct+'%';
      $('undoc').textContent=payload.gap.undocumentedFunctions.length;

      const recs=$('recs'); recs.innerHTML='';
      (payload.gap.recommendations||[]).forEach(r=>{
        const li=document.createElement('li'); li.textContent=r.message; recs.appendChild(li);
      });

      renderChart(payload.history||[]);
      const tb=$('tblUndoc').querySelector('tbody'); tb.innerHTML='';
      (payload.gap.undocumentedFunctions||[]).slice(0,200).forEach(u=>{
        const tr=document.createElement('tr'); const td=document.createElement('td'); td.textContent=u.function; tr.appendChild(td); tb.appendChild(tr);
      });
    }).__ae_dashboardData__();
  }
  function runNow(){
    setStatus('Kjører analyse...');
    google.script.run.withSuccessHandler(function(){ setStatus('Ferdig'); loadLatest(); })
      .ae_runAndLogAnalysis_({});
  }
  function sortUndoc(col){ /* tiny no-op sorter for future columns */ }
  // bootstrap
  loadLatest();
</script>
</body></html>`;
}

// Data provider for dashboard
function __ae_dashboardData__(){
  var latest = ae_runPreview_(); // non-logging quick run if needed
  var hist = __ae_getHistory_();
  return latest ? { health: latest.health, gap: latest.gap, history: hist } : null;
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

/* ------------------------------------ CI helpers ------------------------------------ */
// Minimal JSON you can emit inside CI (clasp or WebApp GET to /exec).
function ae_ciPrintLatest_(){
  var m=ae_getLatestMetrics_(); Logger.log(JSON.stringify(m));
}

/* ---------------------------------- Smoke runner ----------------------------------- */
function runCoreAnalysis_Smoke(){
  var out=ae_runAndLogAnalysis_({});
  __ae_log_().info('runCoreAnalysis_Smoke','ok',{row:out.logRow,health:out.health});
}

/* -------------------------- Notes for CI/CD & Git integration -----------------------
- Add a script property RSP_COMMIT with your current git SHA (or pass {commit:'<sha>'} to ae_runAndLogAnalysis_).
- In GitHub Actions (clasp), run a simple Apps Script function (via clasp run or WebApp) that calls:
      ae_runAndLogAnalysis_({ commit: process.env.GITHUB_SHA });
  Then call ae_ciPrintLatest_() and parse the JSON to decide pass/fail
  (e.g., fail if undocumented increased or score decreased).
------------------------------------------------------------------------------------- */
