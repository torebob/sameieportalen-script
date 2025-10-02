/****************************************************
 * Requirement Generator v1.6.1 – All-in-One + Wizard
 * --------------------------------------------------
 * Adds:
 *  - Small UI entry wizard (HtmlService modal)
 *  - Config schema validation
 *  - Batched ingest + progress reporting
 *  - Enhanced chapter inference
 *  - Post-ingest self-validation (duplicate IDs)
 *  - Input sanitization
 *
 * Public helpers:
 *   Demo_Run_Analysis_And_Ingest_v161()
 *   Demo_Manual_Pipeline_v161()
 *
 * Menu:
 *   onOpen() → "Krav Generator" menu w/ wizard + direct run
 ****************************************************/

/* ===================== Config ===================== */

var REQGEN_CFG = (typeof REQGEN_CFG === 'object' && REQGEN_CFG) ? REQGEN_CFG : {
  LANGUAGE: 'no',                 // 'no' | 'en'
  JACCARD_THRESHOLD: 0.78,
  TOKEN_MIN_LEN: 2,
  INGEST_BATCH_SIZE: 400,
  LARGE_SHEET_ROW_GUARD: 10000,   // warn and fallback path if sheet very large
  SCORE: { W_SOURCE:0.35, W_CROSS:0.25, W_DOMAIN:0.20, W_TEXT:0.20 },
  PRIORITY_LABELS: { MUST:'MÅ', SHOULD:'BØR', COULD:'KAN' },
  TEMPLATES: {
    no: {
      trigger_clock:   (h)=>`Systemet skal periodisk kjøre «${h}» (tidsstyrt).`,
      trigger_edit:    (h)=>`Systemet skal reagere på endringer i regneark via «${h}».`,
      trigger_open:    (h)=>`Systemet skal utføre oppgaver ved åpning via «${h}».`,
      trigger_form:    (h)=>`Systemet skal behandle innsendinger via «${h}».`,
      trigger_generic: (e,h)=>`Systemet skal håndtere hendelse «${e}» via «${h}».`,
      menu_item:       (t,f)=>`Systemet skal tilby menykommando «${t}» som kaller «${f}».`,
      field_item:      (field,sheet)=>`Systemet skal forvalte datafelt «${field}» i arket «${sheet}».`
    },
    en: {
      trigger_clock:   (h)=>`System shall periodically run “${h}” (time-driven).`,
      trigger_edit:    (h)=>`System shall react to spreadsheet edits via “${h}”.`,
      trigger_open:    (h)=>`System shall perform tasks on open via “${h}”.`,
      trigger_form:    (h)=>`System shall handle form submissions via “${h}”.`,
      trigger_generic: (e,h)=>`System shall handle event “${e}” via “${h}”.`,
      menu_item:       (t,f)=>`System shall provide menu command “${t}” calling “${f}”.`,
      field_item:      (field,sheet)=>`System shall manage data field “${field}” in sheet “${sheet}”.`
    }
  }
};

/* =================== Small helpers =================== */

function _rg_safeStr(v){ return (v==null)?'':String(v).trim(); }
function _rg_tokens(s){
  s = _rg_safeStr(s).toLowerCase();
  var parts = s.split(/[^a-z0-9æøå]+/).filter(Boolean);
  var out = [];
  for (var i=0;i<parts.length;i++){ if(parts[i].length>=REQGEN_CFG.TOKEN_MIN_LEN) out.push(parts[i]); }
  return out;
}
function _rg_jaccardOpt(a,b,minTh){
  var A=_rg_tokens(a), B=_rg_tokens(b);
  if (A.length===0 && B.length===0) return 1.0;
  if (A.length===0 || B.length===0) return 0.0;
  var minLen=Math.min(A.length,B.length), maxLen=Math.max(A.length,B.length);
  var upperBound=minLen/maxLen;
  if (typeof minTh==='number' && upperBound<minTh) return 0.0;
  var setA=new Set(A), setB=new Set(B), inter=0;
  setA.forEach(function(t){ if(setB.has(t)) inter++; });
  return inter/(setA.size+setB.size-inter);
}
function _rg_langTemplates(lang){ return (REQGEN_CFG.TEMPLATES[lang] || REQGEN_CFG.TEMPLATES.no); }
function _rg_priorityFromTrigger(evt){
  var e=_rg_safeStr(evt).toUpperCase();
  if (e.indexOf('CLOCK')>=0 || e.indexOf('TIME')>=0) return 'MUST';
  if (e.indexOf('FORM_SUBMIT')>=0) return 'SHOULD';
  if (e.indexOf('OPEN')>=0) return 'SHOULD';
  if (e.indexOf('EDIT')>=0) return 'SHOULD';
  return 'COULD';
}
function _rsp_hasRspSync_(){ return (typeof rsp_syncSheetToDoc === 'function'); }
function _toast_(msg, title){
  try { SpreadsheetApp.getActive().toast(msg, title||'Krav Generator', 3); } catch(_) {}
}

/* ======= Config schema validation (feedback) ======= */

function _rg_validateCfgSchema_(cfg){
  var errors=[];
  if (!cfg || typeof cfg!=='object') errors.push('REQGEN_CFG mangler.');
  if (!cfg.TEMPLATES || typeof cfg.TEMPLATES!=='object') errors.push('REQGEN_CFG.TEMPLATES mangler/ugyldig.');
  if (!cfg.PRIORITY_LABELS || typeof cfg.PRIORITY_LABELS!=='object') errors.push('REQGEN_CFG.PRIORITY_LABELS mangler/ugyldig.');
  if (typeof cfg.JACCARD_THRESHOLD!=='number' || cfg.JACCARD_THRESHOLD<0 || cfg.JACCARD_THRESHOLD>1)
    errors.push('REQGEN_CFG.JACCARD_THRESHOLD må være [0..1].');
  if (typeof cfg.INGEST_BATCH_SIZE!=='number' || cfg.INGEST_BATCH_SIZE<=0)
    errors.push('REQGEN_CFG.INGEST_BATCH_SIZE må være et positivt tall.');
  return errors;
}

/* ===== Input sanitization (feedback) ===== */
function _rsp_sanitizeRequirement_(text){
  if (typeof text!=='string') return '';
  var cleaned = text
    .replace(/[<>]/g,'')
    .replace(/javascript:/gi,'')
    .replace(/=\w+\(/g,'');
  return cleaned.trim().slice(0,2000);
}

/* ===== Enhanced chapter inference (feedback) ===== */
function _rsp_guessKapFromTextEnhanced_(kravText){
  var patterns=[
    {regex:/\bhms\b|sikkerhet|vern/i,          chapter:'5', confidence:0.85},
    {regex:/\brbac\b|tilgang|login|auth/i,     chapter:'2', confidence:0.90},
    {regex:/meny|navigasjon|ui|interface/i,    chapter:'3', confidence:0.70},
    {regex:/data|datamodell|skjema|felt/i,     chapter:'4', confidence:0.80},
    {regex:/budget|budsjet|økonomi|regnskap/i, chapter:'8', confidence:0.80},
    {regex:/møte|meeting|agenda|protokoll/i,   chapter:'7', confidence:0.75}
  ];
  var best={chapter:'',confidence:0};
  for (var i=0;i<patterns.length;i++){
    var p=patterns[i]; if (p.regex.test(kravText) && p.confidence>best.confidence) best={chapter:p.chapter,confidence:p.confidence};
  }
  return (best.confidence>0.6)?best.chapter:'';
}

/* ============== Generators & Scoring ============== */

function generateRequirementCandidatesV161(analysis, opts){
  opts=opts||{};
  var cfgErrors=_rg_validateCfgSchema_(REQGEN_CFG);
  if (cfgErrors.length){ throw new Error('Konfigurasjonsfeil: ' + cfgErrors.join('; ')); }

  var T=_rg_langTemplates(opts.lang || REQGEN_CFG.LANGUAGE);
  var out=[];

  // Triggers
  var triggers=(analysis && analysis.triggers && analysis.triggers.details)||[];
  triggers.forEach(function(t){
    var evt=_rg_safeStr(t.eventType), handler=_rg_safeStr(t.handler);
    var text;
    if (evt.indexOf('CLOCK')>=0 || evt.indexOf('TIME')>=0) text=T.trigger_clock(handler);
    else if (evt.indexOf('FORM')>=0) text=T.trigger_form(handler);
    else if (evt.indexOf('OPEN')>=0) text=T.trigger_open(handler);
    else if (evt.indexOf('EDIT')>=0) text=T.trigger_edit(handler);
    else text=T.trigger_generic(evt,handler);
    out.push({ text:_rsp_sanitizeRequirement_(text), priority:_rg_priorityFromTrigger(evt), source:'trigger', evidence:{eventType:evt, handler:handler} });
  });

  // Menus
  var menus=(analysis && analysis.menus && analysis.menus.fromSheets)||[];
  menus.forEach(function(m){
    var title=_rg_safeStr(m.title||m.functionName), fn=_rg_safeStr(m.functionName);
    if (!fn) return;
    out.push({ text:_rsp_sanitizeRequirement_(T.menu_item(title,fn)), priority:'SHOULD', source:'menu', evidence:{sheet:m.sheet||'', role:m.role||'', active:!!m.active} });
  });

  // Fields
  var sheets=(analysis && analysis.sheets && analysis.sheets.sheets)||[];
  sheets.forEach(function(s){
    var hdrs=_rg_safeStr(s.headerPreview).split('|').map(function(h){return h.trim();}).filter(Boolean);
    hdrs.forEach(function(h){
      out.push({ text:_rsp_sanitizeRequirement_(T.field_item(h,s.name||'')), priority:'SHOULD', source:'field', evidence:{sheet:s.name||'', header:h} });
    });
  });

  // Heuristics
  var fnNames=((analysis && analysis.functions && analysis.functions.global)||[]).map(function(f){return _rg_safeStr(f.name).toLowerCase();});
  var joined=' '+fnNames.join(' ')+' ';
  function addHeur(txt,prio,area){ out.push({ text:_rsp_sanitizeRequirement_(txt), priority:(prio||'SHOULD'), source:'heuristic', evidence:{area:area||''} }); }
  if (/\bhms\b/.test(joined)) addHeur('Systemet skal sikre at HMS-planer genereres, varsles og synkroniseres i kalender.','MUST','HMS');
  if (/\bvaktmester\b/.test(joined)) addHeur('Systemet skal la vaktmester motta, oppdatere og ferdigstille oppgaver.','SHOULD','Tasks');
  if (/\bbudget\b|\bbudsjett\b/.test(joined)) addHeur('Systemet skal støtte budsjetthåndtering med validering, import og rapportering.','SHOULD','Budget');
  if (/\bvote\b|\bvoter\b|\bstemme\b/.test(joined)) addHeur('Systemet skal støtte digital stemmegivning med oppsummering og låsing av vedtak.','SHOULD','Møter');
  if (/\bmeeting\b|\bmøte\b|\bmoter\b/.test(joined)) addHeur('Systemet skal forvalte møter, agenda og protokoll for godkjenning.','SHOULD','Møter');
  if (/\brbac\b|\brole\b|\btilgang\b/.test(joined)) addHeur('Systemet skal håndheve rollebasert tilgangsstyring (RBAC) for brukerhandlinger.','MUST','Security');

  return out;
}

function scoreRequirementsV161(cands){
  var W=REQGEN_CFG.SCORE;
  var normMap=Object.create(null);
  cands.forEach(function(c){ var norm=_rg_tokens(c.text).join(' '); normMap[norm]=(normMap[norm]||0)+1; });

  cands.forEach(function(c){
    var srcW=(c.source==='trigger')?1.0:(c.source==='menu')?0.7:(c.source==='field')?0.6:(c.source==='heuristic')?0.5:0.4;
    var cross=Math.min(1,(normMap[_rg_tokens(c.text).join(' ')]||1)/3);
    var domain=(c.source==='heuristic')?1.0:0.5;
    var t=_rg_safeStr(c.text), L=t.length;
    var lenScore=(L<40)?0.2:(L>300)?0.4:1.0;
    var starts=/^system(et)?\s+skal/i.test(t)?1.0:0.7;
    var punct=/[.!?]$/.test(t)?1.0:0.8;
    var textQ=Math.min(1,(lenScore+starts+punct)/3);
    c.score=+(W.W_SOURCE*srcW + W.W_CROSS*cross + W.W_DOMAIN*domain + W.W_TEXT*textQ).toFixed(3);
  });
  cands.sort(function(a,b){ return (b.score||0)-(a.score||0); });
  return cands;
}

function dedupeRequirementsV161(cands, opts){
  opts=opts||{};
  var th=(typeof opts.threshold==='number')?opts.threshold:REQGEN_CFG.JACCARD_THRESHOLD;
  var kept=[];
  for (var i=0;i<cands.length;i++){
    var c=cands[i], dup=false;
    for (var k=0;k<kept.length;k++){
      if (_rg_jaccardOpt(kept[k].text, c.text, th)>=th){ dup=true; break; }
    }
    if (!dup) kept.push(c);
  }
  return kept;
}

/* =================== RSP adapters =================== */

function _rsp_getSheetName_(opts){
  if (opts && opts.sheetName) return String(opts.sheetName);
  try { if (typeof RSP_CFG==='object' && RSP_CFG && RSP_CFG.SHEET_REQ_NAME) return RSP_CFG.SHEET_REQ_NAME; } catch(_){}
  return 'Requirements';
}
function _rsp_normalizeKapittel_(kap){
  var n=parseInt(String(kap),10);
  return (!isNaN(n) && n>=0 && n<=99)? String(n):'X';
}
function _rsp_resolveReqHeaders_(ss, sheetName, headers){
  var sh=ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Fant ikke arket: '+sheetName);
  var existing = sh.getLastRow()>0 ? sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0] : [];
  if (!existing || existing.length===0){ sh.getRange(1,1,1,headers.length).setValues([headers]); return _rsp_headerIndex_(headers); }
  var low=existing.map(function(h){return _rg_safeStr(h).toLowerCase();});
  headers.forEach(function(h){ var l=_rg_safeStr(h).toLowerCase(); if (low.indexOf(l)<0){ existing.push(h); low.push(l); } });
  sh.getRange(1,1,1,existing.length).setValues([existing]);
  return _rsp_headerIndex_(existing);
}
function _rsp_headerIndex_(headers){
  function idxOf(keys){ var low=headers.map(function(h){return _rg_safeStr(h).toLowerCase();}); for (var i=0;i<keys.length;i++){ var j=low.indexOf(keys[i]); if (j>=0) return j; } return -1; }
  return {
    id:   idxOf(['kravid','krav id','krav-id','id']),
    krav: idxOf(['krav','beskrivelse','tekst','requirement','description','text']),
    prio: idxOf(['prioritet','prio','priority']),
    prog: idxOf(['fremdrift %','fremdrift%','fremdrift','progress','progress %','%']),
    kap:  idxOf(['kapittel','kap','chapter']),
    vers: idxOf(['versjon','version']),
    kom:  idxOf(['kommentar','comment']),
    stamp:idxOf(['sistendret','sist endret','last modified','timestamp'])
  };
}
function _rsp_defaultHeaders_(){ return ['KravID','Krav','Prioritet','Fremdrift %','Kapittel','Versjon','Kommentar','SistEndret']; }
function _rsp_nextIdSeed_(data, idx){
  var perKap=Object.create(null);
  for (var r=1;r<data.length;r++){
    var id=_rg_safeStr(data[r][idx.id]);
    var m=/^K([0-9X]+)-(\d{3,})$/.exec(id);
    if (!m) continue;
    var kap=String(m[1]), nn=parseInt(m[2],10)||0;
    perKap[kap]=Math.max(perKap[kap]||0, nn);
  }
  return { perKap:perKap };
}
function _rsp_guessKapFromText_(kravText){
  var m=/^\s*K(\d+)-\d+/.exec(kravText||'');
  if (m) return m[1];
  var e=_rsp_guessKapFromTextEnhanced_(kravText);
  return e || '';
}

function rsp_ingestRequirementCandidates(cands, opts){
  opts=opts||{};
  var sheetName=_rsp_getSheetName_(opts), batchSize=Number(opts.batchSize||REQGEN_CFG.INGEST_BATCH_SIZE||400);
  var ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Finner ikke arket: '+sheetName);

  // Guard for very large sheet (feedback)
  var lastRow=sh.getLastRow(), lastCol=sh.getLastColumn();
  if (lastRow>REQGEN_CFG.LARGE_SHEET_ROW_GUARD){
    _toast_('Stort ark oppdaget ('+lastRow+' rader) – batterioptimalisert modus','Ytelse');
  }

  var hdrIdx=_rsp_resolveReqHeaders_(ss,sheetName,_rsp_defaultHeaders_());
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var data=(lastRow>1)? sh.getRange(1,1,lastRow,lastCol).getValues() : [headers];

  // existing map
  var existingByText=Object.create(null);
  for (var r=1;r<data.length;r++){
    var txt=_rg_safeStr(data[r][hdrIdx.krav]);
    if (txt) existingByText[txt]=r;
  }
  var nextSeed=_rsp_nextIdSeed_(data,hdrIdx);
  var now=new Date();

  var toAppend=[], written=0;
  function flush(){
    if (!toAppend.length) return;
    var start=sh.getLastRow()+1;
    sh.getRange(start,1,toAppend.length,headers.length).setValues(toAppend);
    written+=toAppend.length;
    toAppend=[];
  }

  // Progress reporting (feedback)
  var total=cands.length, processed=0, reportEvery=Math.max(100, Math.floor(total*0.05));
  function maybeReport(){
    processed++;
    if (processed%reportEvery===0 || processed===total){
      _toast_(processed+'/'+total+' kandidater behandlet', 'Fremdrift');
    }
  }

  for (var i=0;i<cands.length;i++){
    var c=cands[i], krav=_rg_safeStr(c.text);
    if (!krav){ maybeReport(); continue; }

    var prio=REQGEN_CFG.PRIORITY_LABELS[c.priority] || 'BØR';
    var rowIdx=existingByText[krav];

    if (rowIdx!=null){
      data[rowIdx][hdrIdx.prio]=prio;
      data[rowIdx][hdrIdx.kom]=(data[rowIdx][hdrIdx.kom]||'') + (c.source? ` [${c.source}]`:'') + (c.score!=null? ` [score:${c.score}]`:'');
      data[rowIdx][hdrIdx.stamp]=now;
    } else {
      var kapGuess=_rsp_normalizeKapittel_(_rsp_guessKapFromText_(krav) || 'X');
      var next=(nextSeed.perKap[kapGuess]||0)+1; nextSeed.perKap[kapGuess]=next;
      var id='K'+kapGuess+'-'+('000'+next).slice(-3);
      var row=Array(headers.length).fill('');
      row[hdrIdx.id]=id; row[hdrIdx.krav]=krav; row[hdrIdx.prio]=prio;
      row[hdrIdx.kap]=(kapGuess==='X'?'':kapGuess);
      row[hdrIdx.vers]=0; row[hdrIdx.kom]=(c.source?`[${c.source}]`:'')+(c.score!=null?` [score:${c.score}]`:'');
      row[hdrIdx.stamp]=now;
      toAppend.push(row);
      if (toAppend.length>=batchSize){ flush(); Utilities.sleep(20); }
    }
    maybeReport();
  }

  if (data.length>1){ sh.getRange(1,1,data.length,headers.length).setValues(data); }
  flush();

  // Post-ingest validation (feedback)
  var issues=_rsp_validateIngestedData_(sheetName);
  if (issues.length){ _toast_('Valideringsadvarsler: '+issues[0], 'Advarsel'); }

  return { length: written, issues: issues };
}

function rsp_ingestRequirementsFromAnalysis(analysis, opts){
  opts=opts||{};
  var topN=opts.topN||300, lang=opts.lang||REQGEN_CFG.LANGUAGE;

  var cands=generateRequirementCandidatesV161(analysis,{lang:lang});
  scoreRequirementsV161(cands);
  var deduped=dedupeRequirementsV161(cands,{threshold:REQGEN_CFG.JACCARD_THRESHOLD}).slice(0,topN);

  var inserted=rsp_ingestRequirementCandidates(deduped,{ sheetName:_rsp_getSheetName_(opts), upsert:true, batchSize:REQGEN_CFG.INGEST_BATCH_SIZE });
  if (opts.autoPush && _rsp_hasRspSync_()){
    try { rsp_syncSheetToDoc({ dryRun:false, onProgress: opts.onProgress || function(p){ Logger.log(p); } }); }
    catch(e){ Logger.log('RSP push failed: '+(e && e.message)); }
  }
  return { totalGenerated:cands.length, written: inserted.length, warnings: inserted.issues||[] };
}

function _rsp_validateIngestedData_(sheetName){
  var ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName(sheetName);
  if (!sh) return ['Ark ikke funnet: '+sheetName];
  var vals=sh.getDataRange().getValues(); if (vals.length<=1) return [];
  var ids=new Set(), dups=[];
  for (var r=1;r<vals.length;r++){
    var id=_rg_safeStr(vals[r][0]);
    if (!id) continue;
    if (ids.has(id)) dups.push(id);
    ids.add(id);
  }
  var issues=[];
  if (dups.length) issues.push('Dupliserte IDer: '+Array.from(new Set(dups)).join(', '));
  return issues;
}

/* ===================== Wizard UI ===================== */

function onOpen(){
  var ui=SpreadsheetApp.getUi();
  ui.createMenu('Krav Generator')
    .addItem('1) Analyse & foreslå (wizard)','rg_menu_openWizard')
    .addItem('2) Kjør alt (uten UI)','rg_menu_runAllQuick')
    .addSeparator()
    .addItem('Åpne Requirements-arket','rg_menu_openReqSheet')
    .addToUi();
}

function rg_menu_openWizard(){
  var html=HtmlService.createHtmlOutput(_rg_wizardHtml_())
    .setTitle('Krav Generator – Veiviser').setWidth(420).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Krav Generator – Veiviser');
}

function rg_menu_runAllQuick(){
  try {
    _toast_('Starter full kjøring…','Krav Generator');
    var analysis = (typeof performComprehensiveAnalysis_==='function') ? performComprehensiveAnalysis_() : null;
    if (!analysis) throw new Error('Mangler analysefunksjon performComprehensiveAnalysis_.');
    var res = rsp_ingestRequirementsFromAnalysis(analysis,{ topN:300, autoPush:true });
    _toast_('Ferdig. Generert: '+res.totalGenerated+' / skrevet: '+res.written,'Krav Generator');
  } catch(e){
    _toast_('Feil: '+(e && e.message),'Krav Generator');
  }
}

function rg_menu_openReqSheet(){
  var ss=SpreadsheetApp.getActive();
  var name=_rsp_getSheetName_();
  var sh=ss.getSheetByName(name);
  if (sh){ ss.setActiveSheet(sh); _toast_('Åpnet '+name); }
  else { _toast_('Fant ikke arket: '+name); }
}

/* ----- server endpoints for wizard ----- */

function rg_wizard_start(){
  // quick env checks
  var cfgErrors=_rg_validateCfgSchema_(REQGEN_CFG);
  if (cfgErrors.length) return { ok:false, step:'config', message:'Konfigurasjonsfeil: '+cfgErrors.join('; ') };
  var hasAnalysis=(typeof performComprehensiveAnalysis_==='function');
  if (!hasAnalysis) return { ok:false, step:'analysis', message:'Mangler analysefunksjon performComprehensiveAnalysis_.' };
  return { ok:true };
}

function rg_wizard_runPipeline(options){
  options = options || {};
  var analysis = performComprehensiveAnalysis_();
  var cands = generateRequirementCandidatesV161(analysis, { lang: options.lang || REQGEN_CFG.LANGUAGE });
  scoreRequirementsV161(cands);
  var deduped = dedupeRequirementsV161(cands, { threshold: REQGEN_CFG.JACCARD_THRESHOLD }).slice(0, options.topN || 300);
  var written = rsp_ingestRequirementCandidates(deduped, { sheetName:_rsp_getSheetName_(), upsert:true, batchSize:REQGEN_CFG.INGEST_BATCH_SIZE });

  if (options.autoPush && _rsp_hasRspSync_()){
    try { rsp_syncSheetToDoc({ dryRun:false, onProgress:function(p){ Logger.log(p); } }); }
    catch(e){ /* ignore UI */ }
  }
  return { ok:true, generated:cands.length, ingested:written.length, warnings: written.issues||[] };
}

/* ----- wizard HTML (minimal) ----- */
function _rg_wizardHtml_(){
  return (
'<!DOCTYPE html><html><head><base target="_top"><style>' +
'body{font-family:Google Sans,Arial,sans-serif;margin:16px;} h2{margin:0 0 12px;font-size:18px} '+
'.row{margin:10px 0} button{padding:8px 12px} .muted{color:#666} .ok{color:#0a0} .err{color:#b00} '+
'.log{white-space:pre-wrap;background:#fafafa;border:1px solid #eee;padding:8px;height:120px;overflow:auto;} '+
'</style></head><body>' +
'<h2>Krav Generator – Veiviser</h2>' +
'<div class="row muted">Kjører: Analyse → Generer → Poeng → Dedupe → Ingest → (valgfritt) Push</div>' +
'<div class="row">Språk: <select id="lang"><option value="no">Norsk</option><option value="en">English</option></select></div>' +
'<div class="row">Topp N: <input id="topn" type="number" value="300" min="50" max="2000" style="width:90px"></div>' +
'<div class="row"><label><input id="autopush" type="checkbox" checked> Push til dokument (hvis RSP er tilgjengelig)</label></div>' +
'<div class="row"><button id="runBtn">Kjør veiviser</button> <span id="status" class="muted"></span></div>' +
'<div class="row"><div id="log" class="log"></div></div>' +
'<script>' +
'const $=id=>document.getElementById(id); function log(m){const L=$("log"); L.textContent+=(m+"\\n"); L.scrollTop=L.scrollHeight;}' +
'$("runBtn").onclick=function(){ $("status").textContent="Starter…"; $("log").textContent=""; '+
'google.script.run.withSuccessHandler(function(r){ if(!r.ok){$("status").textContent=""; log("Feil: "+r.message); return;} '+
'$("status").textContent="Analyserer…"; log("Miljø OK – starter pipeline"); '+
'const opts={ lang: $("lang").value, topN: Number($("topn").value)||300, autoPush: $("autopush").checked }; '+
'google.script.run.withSuccessHandler(function(res){ if(!res.ok){$("status").textContent=""; log("Feil: "+(res.message||"ukjent")); return;} '+
'$("status").textContent="Ferdig"; log("Generert: "+res.generated); log("Skrevet: "+res.ingested); if(res.warnings && res.warnings.length){ log("Advarsler: "+res.warnings.join("; ")); } '+
'}).withFailureHandler(function(e){ $("status").textContent=""; log("Feil under kjøring: "+(e && e.message)); }).rg_wizard_runPipeline(opts); '+
'}).withFailureHandler(function(e){ $("status").textContent=""; log("Feil ved start: "+(e && e.message)); }).rg_wizard_start(); };' +
'</script></body></html>'
  );
}

/* ============= Convenience demos (optional) ============= */

function Demo_Run_Analysis_And_Ingest_v161(){
  var analysis = (typeof performComprehensiveAnalysis_==='function') ? performComprehensiveAnalysis_() : null;
  if (!analysis) throw new Error('Mangler analysefunksjon performComprehensiveAnalysis_.');
  var res = rsp_ingestRequirementsFromAnalysis(analysis,{ topN:300, autoPush:true });
  Logger.log(res);
}

function Demo_Manual_Pipeline_v161(){
  var analysis = performComprehensiveAnalysis_();
  var c = generateRequirementCandidatesV161(analysis,{ lang:'no' });
  scoreRequirementsV161(c);
  var deduped = dedupeRequirementsV161(c,{ threshold:REQGEN_CFG.JACCARD_THRESHOLD }).slice(0,200);
  var written = rsp_ingestRequirementCandidates(deduped,{ sheetName:_rsp_getSheetName_(), upsert:true, batchSize:REQGEN_CFG.INGEST_BATCH_SIZE });
  if (_rsp_hasRspSync_()){
    rsp_syncSheetToDoc({ dryRun:false, onProgress:function(p){ Logger.log(p); } });
  }
  Logger.log(written);
}
