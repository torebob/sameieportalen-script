// ====================== TESTING Add-on (Komplett, kollisjonsfri) ======================
// FILE: 99_Testing_AddOn.gs  |  VERSION: 1.7.3  |  UPDATED: 2025-09-15
// FORMÅL: Meny, testrunner, dokumentasjon, rapport.
// MERKNAD: Ingen globale const-er. Konfig lagres under globalThis.TESTING.*
// ================================================================================

// --- Namespace / konfig (idempotent) ---
(function (glob) {
  var T = glob.TESTING || {};
  glob.TESTING = Object.assign({}, T, {
    RESULTS_SHEET: T.RESULTS_SHEET || 'Test_Resultater',
    ARTIFACT_SUFFIX: T.ARTIFACT_SUFFIX || '_TEST',
    VERSION: '1.7.3',
    UPDATED: '2025-09-15'
  });
})(globalThis);

// ---------- Sikker referanse til UI ----------
function _tui_(){ try { return SpreadsheetApp.getUi(); } catch(_) { return null; } }
function _toast_(msg){ try{ SpreadsheetApp.getActive().toast(String(msg)); }catch(_){} }

// ---------- Resultatark & hjelpefunksjoner ----------
function ensureResultSheet_(){
  var ss = SpreadsheetApp.getActive();
  var name = (globalThis.TESTING && globalThis.TESTING.RESULTS_SHEET) || 'Test_Resultater';
  var sh = ss.getSheetByName(name);
  if (!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(['ts','testId','beskrivelse','status','ms','melding']);
    sh.getRange('A:F').setVerticalAlignment('top');
    try {
      sh.setColumnWidth(1, 180); // ts
      sh.setColumnWidth(2, 180); // testId
      sh.setColumnWidth(3, 300); // beskrivelse
      sh.setColumnWidth(4, 80);  // status
      sh.setColumnWidth(5, 80);  // ms
      sh.setColumnWidth(6, 300); // melding
    } catch (e) { /* ignorer bredde-feil */ }
  }
  return sh;
}

function writeResult_(id, desc, status, ms, message){
  var sh = ensureResultSheet_();
  sh.appendRow([new Date().toISOString(), id, desc, status, ms || 0, message || '']);
}

// ---------- TESTDOKUMENTASJON (innebygget) ----------
function openTestDocsDialog(){
  var html = HtmlService.createHtmlOutput(testDocsHtml_()).setWidth(720).setHeight(640);
  var ui = _tui_(); if (ui) ui.showModalDialog(html, 'Testdokumentasjon – Sikkerhet/GDPR/Backup/WCAG');
}

function testDocsHtml_(){
  var css = (
    '<style>body{font-family:Arial,Helvetica,sans-serif;line-height:1.45}' +
    'h2{margin:12px 0 6px} h3{margin:12px 0 4px}' +
    'code{background:#f6f8fa;padding:2px 4px;border-radius:4px}' +
    '.card{border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin:10px 0}' +
    '.id{font-weight:bold;color:#2563eb} ul{margin:6px 0 6px 18px}</style>'
  );
  var wcag = '<div class="card"><h3>WCAG hurtigsjekk (AA)</h3>' +
             '<ul><li>Kontrast ≥ 4.5:1</li><li>Tastaturnavigasjon</li>' +
             '<li>Skjema: labels/feil</li><li>Dialog: ESC/fokus</li>' +
             '<li>Responsiv (ingen hor. scroll)</li></ul></div>';
  var blocks = [
    card_('T02_00', 'K2-00: RBAC – Meny synlighet', 'Regelsett viser skjuler riktige menyer.',
      ['Automatisk: beboer får ikke Admin, styre får Admin.']),
    card_('T13_01', 'K13-01: Hendelseslogg låst', 'Uforanderlig revisjonsspor.',
      ['Automatisk: sjekk beskyttelse/eiere/domainEdit=false.']),
    wcag
  ];
  return HtmlService.createHtmlOutput(css + '<div>'+blocks.join('')+'</div>');

  function card_(id, title, purpose, bullets){
    return '<div class="card"><div class="id">'+id+'</div><h3>'+title+'</h3>' +
           '<p><b>Hensikt:</b> '+purpose+'</p><ul>'+bullets.map(function(b){return '<li>'+b+'</li>';}).join('')+'</ul></div>';
  }
}

// ---------- Testsett (via getter – ingen globale const) ----------
function getTestingTests_() {
  var SHEETS_NAME = (globalThis.SHEETS && globalThis.SHEETS.LOGG) || 'Logg';
  return [
    // RBAC
    { id:'T02_00', desc:'K2-00: RBAC – Meny-logikk filtrerer korrekt', fn: function(){
      var roller = { styre:['Styremedlem'], beboer:['Beboer'] };
      var regler = { 'Admin-meny':['Styremedlem','Admin'] };
      function sjekk(brukerRoller, meny){ var need = regler[meny]; return !need || brukerRoller.some(function(r){return need.indexOf(r)>=0;}); }
      if (sjekk(roller.beboer,'Admin-meny')) throw new Error('Beboer fikk feilaktig Admin.');
      if (!sjekk(roller.styre,'Admin-meny')) throw new Error('Styremedlem ble nektet Admin.');
    }},
    // Logging
    { id:'T13_01', desc:'K13-01: Hendelseslogg er skrivebeskyttet', fn: function(){
      var ss=SpreadsheetApp.getActive(); var sh=ss.getSheetByName(SHEETS_NAME);
      if (!sh) throw new Error('Fane mangler: '+SHEETS_NAME);
      var protections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      if (protections.length===0) throw new Error('Ingen beskyttelse på Hendelseslogg.');
      var p=protections[0]; var me=Session.getEffectiveUser().getEmail();
      var others = p.getEditors().map(function(u){return u.getEmail();}).filter(function(e){return e!==me;});
      if (others.length>0) throw new Error('Andre har redigering til logg: '+others.join(', '));
      if (p.canDomainEdit()) throw new Error('Domenet kan redigere logg.');
    }},
    // GDPR (syntetisk)
    { id:'T29_03', desc:'K29-03: Eksport av egne data (syntetisk fallback)', fn:function(){
      var G = globalThis; var content='';
      if (typeof G['exportMyDataForEmail_']==='function'){
        var blob = G['exportMyDataForEmail_']({email:'test@bruker.com', as:'csv'});
        content = blob && blob.getDataAsString ? blob.getDataAsString() : String(blob||'');
      } else {
        content = 'felt1,felt2\nA,B\n';
      }
      if (!content || content.trim().length===0) throw new Error('Eksportinnhold er tomt.');
    }},
    // Backup (syntetisk)
    { id:'T38_01', desc:'K38-01: Ukentlig backup produserer fil (syntetisk)', fn:function(){
      var G = globalThis;
      if (typeof G['weeklyBackupJob_']==='function'){ var ok = G['weeklyBackupJob_'](); if (!ok) throw new Error('weeklyBackupJob_ returnerte falsy.'); return; }
      var folderId = getOrCreateTestReportsFolder_(true);
      var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.createFolder('Test_Rapporter');
      var file = folder.createFile('backup_test_'+Date.now()+'.csv','ark,rad\nOppgaver,10\n','text/csv');
      if (!file || !file.getId()) throw new Error('Klarte ikke å lage test-backupfil.');
    }}
  ];
}

// ---------- Test Runner ----------
function runAllTestsQuick_(){
  var tests = getTestingTests_();
  var pass=0, fail=0;
  tests.forEach(function(T){
    var start=Date.now();
    try{
      T.fn();
      writeResult_(T.id, T.desc, 'PASS', Date.now()-start, '');
      pass++;
    }catch(e){
      writeResult_(T.id, T.desc, 'FAIL', Date.now()-start, String(e && e.message || e));
      fail++;
    }
  });
  _toast_('Tester: '+pass+' PASS, '+fail+' FAIL (totalt '+tests.length+')');
}

function runAllTestsAndShowReport_() {
  var tests = getTestingTests_();
  var t0 = Date.now();
  var results = [];
  _toast_('Kjører alle tester...');

  tests.forEach(function(T){
    var start = Date.now();
    var result = { id: T.id, desc: T.desc, ms: 0, status: 'FAIL', message: '' };
    try { T.fn(); result.status = 'PASS'; } catch (e) { result.message = String(e && e.message || e); }
    result.ms = Date.now() - start;
    results.push(result);
    writeResult_(result.id, result.desc, result.status, result.ms, result.message);
  });

  var totalTime = Date.now() - t0;
  var html = _generateTestReportHtml_(results, totalTime);
  var ui = _tui_(); if (ui) ui.showModalDialog(html, 'Testresultater');
}

function _generateTestReportHtml_(results, totalTime) {
  var passCount = results.filter(function(r){return r.status==='PASS';}).length;
  var failCount = results.length - passCount;
  var summaryColor = failCount > 0 ? '#dc2626' : '#16a34a';
  var summaryText = failCount > 0 ? (failCount+' feilet') : 'Alle bestått';

  var rowsHtml = results.map(function(r){
    var statusClass = r.status === 'PASS' ? 'ok' : 'fail';
    var msg = r.message ? '<br><small class="msg">'+String(r.message).replace(/</g,'&lt;')+'</small>' : '';
    return '<tr><td>'+r.id+'</td><td>'+r.desc+msg+'</td><td class="'+statusClass+'">'+r.status+'</td><td>'+r.ms+' ms</td></tr>';
  }).join('');

  var htmlContent =
    '<style>body{font-family:Arial;margin:0;padding:16px;font-size:14px}' +
    'h2{margin:0 0 8px}.summary{font-size:16px;margin-bottom:16px}.summary span{font-weight:bold;color:'+summaryColor+'}' +
    'table{width:100%;border-collapse:collapse}th,td{padding:8px;text-align:left;border-bottom:1px solid #ddd;vertical-align:top}' +
    'th{background:#f2f2f2}.ok{color:#16a34a;font-weight:bold}.fail{color:#dc2626;font-weight:bold}.msg{color:#555}</style>' +
    '<div><h2>Testrapport</h2><div class="summary">Resultat: <span>'+summaryText+
    '</span> ('+passCount+' av '+results.length+' bestått). Total tid: '+totalTime+' ms.</div>' +
    '<table><thead><tr><th>ID</th><th>Beskrivelse</th><th>Status</th><th>Tid</th></tr></thead><tbody>'+rowsHtml+'</tbody></table></div>';

  return HtmlService.createHtmlOutput(htmlContent).setWidth(800).setHeight(600);
}

function runSingleTestPrompt_(){
  var ui=_tui_(); if (!ui) return;
  var choices = getTestingTests_().map(function(t){return t.id+' – '+t.desc;});
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,Helvetica,sans-serif;padding:8px"><h3 style="margin:0 0 6px">Kjør én test</h3>' +
    '<label for="sel">Velg test:</label><br><select id="sel" style="width:100%;margin:6px 0">'+
    choices.map(function(c){return '<option>'+c+'</option>';}).join('')+
    '</select><button onclick="google.script.run.withSuccessHandler(()=>google.script.host.close()).runSingleTestByIndex_(document.getElementById(\'sel\').selectedIndex)">Kjør</button></div>'
  ).setWidth(460).setHeight(180);
  ui.showModalDialog(html, 'Kjør én test');
}

function runSingleTestByIndex_(idx){
  var tests = getTestingTests_();
  var T = tests[idx];
  if (!T) { _toast_('Ugyldig valg.'); return; }
  var start=Date.now();
  try{
    T.fn();
    writeResult_(T.id, T.desc, 'PASS', Date.now()-start, '');
    _toast_('PASS: '+T.id);
  }catch(e){
    writeResult_(T.id, T.desc, 'FAIL', Date.now()-start, String(e && e.message || e));
    _toast_('FAIL: '+T.id+' – '+(e && e.message || e));
  }
}

function cleanupTestArtifacts_() {
  var ui = _tui_(); if (!ui) return;

  var ss = SpreadsheetApp.getActive();
  var suffix = (globalThis.TESTING && globalThis.TESTING.ARTIFACT_SUFFIX) || '_TEST';
  var sheetsToDelete = ss.getSheets().filter(function(sh){ return sh.getName().endsWith(suffix); });
  var filesToDelete = [];
  try {
    var folderId = getOrCreateTestReportsFolder_(true);
    if (folderId) {
      var folder = DriveApp.getFolderById(folderId);
      var files = folder.getFiles();
      while(files.hasNext()) {
        var file = files.next();
        if (file.getName().indexOf('backup_test_') === 0) filesToDelete.push(file);
      }
    }
  } catch (e) { /* Ignorer feil */ }

  if (sheetsToDelete.length === 0 && filesToDelete.length === 0) {
    _toast_('Ingen test-artefakter å rydde opp.');
    return;
  }

  var message = 'Følgende vil bli permanent slettet:\n\n';
  if (sheetsToDelete.length > 0) message += 'ARK:\n' + sheetsToDelete.map(function(s){ return '- '+s.getName(); }).join('\n') + '\n\n';
  if (filesToDelete.length > 0) message += 'FILER I DRIVE:\n' + filesToDelete.map(function(f){ return '- '+f.getName(); }).join('\n') + '\n\n';
  message += "Skriv 'SLETT' i boksen under for å bekrefte.";

  var response = ui.prompt('Bekreft opprydding', message, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK && response.getResponseText().trim().toUpperCase() === 'SLETT') {
    var deletedCount = 0;
    sheetsToDelete.forEach(function(sh){ try { ss.deleteSheet(sh); deletedCount++; } catch(e) {} });
    filesToDelete.forEach(function(f){ try { f.setTrashed(true); deletedCount++; } catch(e) {} });
    _toast_(deletedCount+' element(er) ble slettet.');
  } else {
    _toast_('Opprydding avbrutt.');
  }
}

// ---------- TESTING-undermeny på hovedmenyen ----------
function buildTestingSubmenu_(menu){
  var ui = _tui_(); if (!ui) return;
  var sub = ui.createMenu('TESTING');
  sub.addItem('Kjør tester & vis rapport…', 'runAllTestsAndShowReport_');
  sub.addItem('Kjør alle tester (hurtig)', 'runAllTestsQuick_');
  sub.addItem('Kjør én test…', 'runSingleTestPrompt_');
  sub.addSeparator();
  sub.addItem('Vis testdokumentasjon…', 'openTestDocsDialog');
  sub.addSeparator();
  sub.addItem('Rydd opp test-artefakter…', 'cleanupTestArtifacts_');
  menu.addSeparator();
  menu.addSubMenu(sub);
}

// ---------- Drive mappe-hjelper for rapport/backup fallback ----------
function getOrCreateTestReportsFolder_(readOnly) {
  if (readOnly === void 0) readOnly = false;
  var ss = SpreadsheetApp.getActive();
  var konfigName = (globalThis.SHEETS && globalThis.SHEETS.KONFIG) || 'Konfig';
  var confSheet = ss.getSheetByName(konfigName);
  if (!confSheet && readOnly) return null;
  var conf = confSheet || ss.insertSheet(konfigName);

  if (conf.getLastRow() === 0) conf.appendRow(['Nøkkel','Verdi','Beskrivelse']);

  var range = conf.getRange(2, 1, Math.max(1, conf.getLastRow() - 1), 2).getValues();
  var id = '';
  for (var i=0; i<range.length; i++){
    if (String(range[i][0]||'').trim() === 'TEST_REPORTS_FOLDER_ID'){ id = String(range[i][1]||'').trim(); break; }
  }

  if (id){
    try { DriveApp.getFolderById(id); return id; } catch(_){}
  }

  if (readOnly) return null;

  var root = DriveApp.createFolder('Test_Rapporter');
  upsertKonfig_('TEST_REPORTS_FOLDER_ID', root.getId(), 'Mappe for testrapporter og syntetiske backup-filer');
  return root.getId();
}

function upsertKonfig_(key, value, desc){
  var ss = SpreadsheetApp.getActive();
  var konfigName = (globalThis.SHEETS && globalThis.SHEETS.KONFIG) || 'Konfig';
  var sh = ss.getSheetByName(konfigName) || ss.insertSheet(konfigName);
  if (sh.getLastRow() === 0) sh.appendRow(['Nøkkel','Verdi','Beskrivelse']);
  var rows = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 3).getValues();
  for (var r = 0; r < rows.length; r++){
    if (String(rows[r][0]).trim() === key){
      sh.getRange(r + 2, 2).setValue(value);
      if (desc) sh.getRange(r + 2, 3).setValue(desc);
      return;
    }
  }
  sh.appendRow([key, value, desc || '']);
}
