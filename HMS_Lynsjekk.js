// =============================================================================
// HMS – Lynsjekk & Hurtigreparasjon
// FILE: 62_HMS_Lynsjekk.gs
// VERSION: 1.0.0
// UPDATED: 2025-09-15
// KREVER: 60_HMS_Vedlikeholdsplan.gs (v1.2.0+)
// OPPGAVER: hmsLynsjekk(), hmsInstallTriggers(), hmsQuickRepair()
// =============================================================================

var DIAG_SHEET = 'DIAG_HMS';
var REQ_PLAN_COLS = [
'PlanID','System','Komponent','Oppgave','Beskrivelse','Frekvens',
'PreferertMåned','NesteStart','AnsvarligRolle','Leverandør','LeverandørKontakt',
'Myndighetskrav','Standard/Referanse','Kritikalitet(1-5)',
'EstTidTimer','EstKost','HistoriskKost','BudsjettKonto',
'DokumentasjonURL','SjekklisteURL','Lokasjon','Byggnummer','Garantistatus',
'SesongAvhengig','SistUtført','Kommentar','Aktiv'
];
var REQ_TASKS_COLS = [
'Tittel','Kategori','Status','Frist','Opprettet','Ansvarlig',
'Seksjonsnr','PlanID','AutoKey','System','Komponent','Lokasjon','Byggnummer',
'Myndighetskrav','Kritikalitet','Hasteprioritering',
'EstKost','BudsjettKonto','FaktiskKost',
'DokumentasjonURL','SjekklisteURL','Garantistatus','BeboerVarsling','Værforhold','Leverandør','LeverandørKontakt',
'Kommentar','OppdatertAv','Oppdatert'
];
var VALID_FREQ = ['MÅNEDLIG','MANEDLIG','KVARTAL','HALVÅR','HALVAR','ÅRLIG','AARLIG','2ÅR','2AAR','3ÅR','3AAR','5ÅR','5AAR','10ÅR','10AAR'];
var VALID_STATUS = ['Åpen','Utført','Avlyst'];
var VALID_PRI = ['Lav','Normal','Høy','Kritisk'];

function hmsLynsjekk() {
var ss = SpreadsheetApp.getActive();
var diag = ss.getSheetByName(DIAG_SHEET) || ss.insertSheet(DIAG_SHEET);
diag.clear(); diag.getRange(1,1,1,3).setValues([['Del','Status','Detaljer']]).setFontWeight('bold');
var row = 2;

function put(sec, ok, detail) {
var mark = ok===true ? '✅' : (ok==='warn' ? '⚠️' : '❌');
diag.getRange(row,1,1,3).setValues([[sec, mark, String(detail||'')]]); row++;
}

// 1) Sjekk ark
var plan = ss.getSheetByName('HMS_PLAN');
var tasks = ss.getSheetByName('TASKS');
var access = ss.getSheetByName('TILGANG');
var beboere = ss.getSheetByName('BEBOERE');
var lever = ss.getSheetByName('LEVERANDØRER');
put('Ark: HMS_PLAN', !!plan, plan ? 'OK' : 'Manglende ark – kjør hmsMigrateSchema_v1_1()');
put('Ark: TASKS', !!tasks, tasks ? 'OK' : 'Manglende ark – kjør hmsMigrateSchema_v1_1()');
put('Ark: TILGANG', !!access, access ? 'OK' : 'Anbefalt ark for roller (Email,Rolle) mangler');
put('Ark: BEBOERE', !!beboere ? 'warn' : false, beboere ? 'OK (valgfritt for e-postvarsling)' : 'Valgfritt – gir beboervarsling');
put('Ark: LEVERANDØRER', !!lever ? 'warn' : 'warn', lever ? 'OK (valgfritt for leverandøroppslag)' : 'Valgfritt – anbefales');

// 2) Kolonnekrav
if (plan) {
var ph = plan.getRange(1,1,1,plan.getLastColumn()).getValues()[0];
var missP = diff(REQ_PLAN_COLS, ph);
put('HMS_PLAN kolonner', missP.length===0 ? true : false, missP.length?('Mangler: '+missP.join(', ')):'OK');
}
if (tasks) {
var th = tasks.getRange(1,1,1,tasks.getLastColumn()).getValues()[0];
var missT = diff(REQ_TASKS_COLS, th);
put('TASKS kolonner', missT.length===0 ? true : false, missT.length?('Mangler: '+missT.join(', ')):'OK');
}

// 3) Datakvalitet i HMS_PLAN
if (plan && plan.getLastRow()>1) {
var pvals = plan.getDataRange().getValues(); pvals.shift();
var pidx = byName(plan.getRange(1,1,1,plan.getLastColumn()).getValues()[0]);
var dupePlan = dupes(pvals.map(function(r){ return String(r[pidx.PlanID-1]||'').trim(); }).filter(String));
put('PlanID unike', dupePlan.length?false:true, dupePlan.length?('Duplikater: '+dupePlan.join(', ')):'OK');

var badFreq = [];
var badDate = [];
var badBygg = [];
var badYN = [];
for (var i=0;i<pvals.length;i++){
  var r = pvals[i];
  var id = String(r[pidx.PlanID-1]||'').trim() || ('rad '+(i+2));
  var f = _normFreq_(String(r[pidx.Frekvens-1]||''));
  if (VALID_FREQ.indexOf(f)<0) badFreq.push(id);
  var ns = r[pidx.NesteStart-1]; if (ns && !(ns instanceof Date) && isNaN(new Date(ns).getTime())) badDate.push(id);
  if (pidx.Byggnummer && r[pidx.Byggnummer-1] && (Number(r[pidx.Byggnummer-1])<1 || Number(r[pidx.Byggnummer-1])>99)) badBygg.push(id);
  var ynFields = ['Myndighetskrav','SesongAvhengig','Aktiv']; 
  for (var y=0;y<ynFields.length;y++){
    var col = pidx[ynFields[y]]; if (!col) continue;
    var v = String(r[col-1]||'').toLowerCase();
    if (v && ['ja','nei','true','false','1','0'].indexOf(v)<0) badYN.push(id+' ('+ynFields[y]+')');
  }
}
put('Frekvens gyldig', badFreq.length? 'warn' : true, badFreq.length?('Kontroller: '+badFreq.join(', ')):'OK');
put('NesteStart dato', badDate.length? 'warn' : true, badDate.length?('Kontroller: '+badDate.join(', ')):'OK');
put('Byggnummer (1-99)', badBygg.length? 'warn' : true, badBygg.length?('Kontroller: '+badBygg.join(', ')):'OK');
put('Ja/Nei-felt', badYN.length? 'warn' : true, badYN.length?('Kontroller: '+badYN.join(', ')):'OK');


}

// 4) Datakvalitet i TASKS
if (tasks && tasks.getLastRow()>1) {
var tvals = tasks.getDataRange().getValues(); tvals.shift();
var tidx = byName(tasks.getRange(1,1,1,tasks.getLastColumn()).getValues()[0]);
var autoKeys = tvals.map(function(r){ return String(r[tidx.AutoKey-1]||'').trim(); }).filter(String);
var dupeAK = dupes(autoKeys);
put('AutoKey unike', dupeAK.length?false:true, dupeAK.length?('Duplikater: '+dupeAK.slice(0,20).join(', ') + (dupeAK.length>20?' ...':'') ):'OK');

var badStatus=[], badPri=[], badDue=[], emptyWeather=[];
var now = new Date();
for (var j=0;j<tvals.length;j++){
  var r2 = tvals[j];
  var id2 = String(r2[tidx.AutoKey-1]||('rad '+(j+2)));
  var st = String(r2[tidx.Status-1]||'');
  if (VALID_STATUS.indexOf(st)<0) badStatus.push(id2);
  var pr = String(r2[tidx.Hasteprioritering-1]||'');
  if (pr && VALID_PRI.indexOf(pr)<0) badPri.push(id2);
  var due = r2[tidx.Frist-1]; if (!(due instanceof Date)) badDue.push(id2);
  var lok = String(r2[tidx.Lokasjon-1]||''); 
  var wf = String(r2[tidx.Værforhold-1]||'');
  if (/ute|utendørs/i.test(lok) && String(r2[tidx.Status-1]||'')==='Åpen' && (!wf || wf==='N/A')) emptyWeather.push(id2);
}
put('Status gyldig', badStatus.length? 'warn' : true, badStatus.length?('Kontroller: '+badStatus.join(', ')):'OK');
put('Prioritet gyldig', badPri.length? 'warn' : true, badPri.length?('Kontroller: '+badPri.join(', ')):'OK');
put('Frist er dato', badDue.length? 'warn' : true, badDue.length?('Kontroller: '+badDue.join(', ')):'OK');
put('Værforhold for ute', emptyWeather.length? 'warn' : true, emptyWeather.length?('Uteoppg. uten værforhold: '+emptyWeather.join(', ')):'OK');


}

// 5) Triggere
var tr = ScriptApp.getProjectTriggers();
var hasGen = tr.some(function(t){ return t.getHandlerFunction()==='hmsGenerateTasks'; });
var hasNotify = tr.some(function(t){ return t.getHandlerFunction()==='hmsNotifyResidents'; });
put('Trigger: hmsGenerateTasks', hasGen ? true : 'warn', hasGen ? 'OK' : 'Anbefalt daglig 03:15');
put('Trigger: hmsNotifyResidents', hasNotify ? 'warn' : 'warn', hasNotify ? 'OK' : 'Anbefalt daglig 07:30');

// 6) Tilgang
try {
var prof = typeof hmsGetUserProfile==='function' ? hmsGetUserProfile() : {email:'',role:'LESER',canEdit:false};
put('Tilgang (din rolle)', prof.canEdit ? true : 'warn', (prof.email||'?')+' • '+(prof.role||'?'));
} catch(e) {
put('Tilgang (din rolle)', 'warn', 'Kunne ikke lese – har du lagt inn 60-filen?');
}

diag.autoResizeColumns(1,3);
return 'Lynsjekk fullført → se ark "'+DIAG_SHEET+'"';
}

function hmsInstallTriggers() {
// Oppretter standard triggere (idempotent)
var existing = ScriptApp.getProjectTriggers().map(function(t){ return t.getHandlerFunction(); });
if (existing.indexOf('hmsGenerateTasks')<0) {
ScriptApp.newTrigger('hmsGenerateTasks').timeBased().atHour(3).nearMinute(15).everyDays(1).create();
}
if (existing.indexOf('hmsNotifyResidents')<0) {
ScriptApp.newTrigger('hmsNotifyResidents').timeBased().atHour(7).nearMinute(30).everyDays(1).create();
}
return 'Triggere satt (03:15 generering, 07:30 varsling).';
}

function hmsQuickRepair() {
// Trygge småreparasjoner: fyll NextStart, normaliser Ja/Nei, frekvens, Aktiv
var ss = SpreadsheetApp.getActive();
var plan = ss.getSheetByName('HMS_PLAN'); if (!plan || plan.getLastRow()<2) return 'Ingen HMS_PLAN-data.';
var ph = plan.getRange(1,1,1,plan.getLastColumn()).getValues()[0]; var pidx = byName(ph);
var vals = plan.getRange(2,1,plan.getLastRow()-1,plan.getLastColumn()).getValues();
var y = new Date().getFullYear(); var fixed=0;

for (var i=0;i<vals.length;i++){
// NesteStart
if (pidx.NesteStart) {
var v = vals[i][pidx.NesteStart-1];
if (!v) { vals[i][pidx.NesteStart-1] = new Date(y,0,1); fixed++; }
}
// Aktiv default Ja
if (pidx.Aktiv) {
var a = String(vals[i][pidx.Aktiv-1]||'').trim().toLowerCase();
if (!a) { vals[i][pidx.Aktiv-1] = 'Ja'; fixed++; }
}
// Ja/Nei normalisering
['Myndighetskrav','SesongAvhengig'].forEach(function(col){
if (pidx[col]) {
var t = String(vals[i][pidx[col]-1]||'').trim().toLowerCase();
if (t==='true'||t==='1') { vals[i][pidx[col]-1]='Ja'; fixed++; }
if (t==='false'||t==='0') { vals[i][pidx[col]-1]='Nei'; fixed++; }
}
});
// Frekvens normalisering
if (pidx.Frekvens) {
var f = normFreq(String(vals[i][pidx.Frekvens-1]||''));
vals[i][pidx.Frekvens-1] = f;
}
}
plan.getRange(2,1,vals.length,vals[0].length).setValues(vals);
return 'Hurtigreparasjon gjort: '+fixed+' felt justert.';
}

// -------------------------- Hjelpefunksjoner --------------------------

function diff(need, have) {
var set = {}; for (var i=0;i<have.length;i++) set[String(have[i]).trim()] = true;
var out=[]; for (var j=0;j<need.length;j++) if (!set[need[j]]) out.push(need[j]);
return out;
}
function dupes(arr) {
var seen={}, out=[]; for (var i=0;i<arr.length;i++){ var k=arr[i]; if (!k) continue; if (seen[k]) { if (out.indexOf(k)<0) out.push(k); } else seen[k]=1; }
return out;
}
function byName(header) { var map={}; for (var i=0;i<header.length;i++){ var h=String(header[i]||'').trim(); if(h) map[h]=i+1; } return map; }
function normFreq(s) {
var f = String(s||'').toUpperCase().replace('Å','A').replace('Ø','O').replace('Æ','AE').trim();
// mapper noen vanlige varianter
if (f==='MND' || f==='MNDLIG') f='MANEDLIG';
if (f==='KVARTALVIS' || f==='KVARTALSVIS') f='KVARTAL';
if (f==='HALVAAR') f='HALVAR';
if (!f) f='AARLIG';
return f;
}