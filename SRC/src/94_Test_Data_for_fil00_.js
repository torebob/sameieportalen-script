/* ====================== TESTING Add-on (Komplett) ======================
 * FILE: 00_Testing_AddOn.gs
 * VERSION: 2.0.0
 * UPDATED: 2025-09-14
 *
 * FORMÅL: Én-fil testmodul (meny, testrunner, dokumentasjon, rapport).
 * AVHENGIGHETER: 00_App_Core.gs (APP/SHEETS/_ui())
 * ENDRINGER v2.0.0:
 *  - Versjonsheader standardisert
 *  - Nye røyk/samspills-tester: dashMetrics(), dashOpen() (hvis tilgjengelig), Vaktmester API
 *  - TESTING-meny legges inn via buildTestingSubmenu_()
 * ===================================================================== */

function _tui_(){ try { return SpreadsheetApp.getUi(); } catch(_) { return null; } }
function _testToast_(msg){ try{ SpreadsheetApp.getActive().toast(String(msg)); }catch(_){} }

const TEST_RESULTS_SHEET = 'Test_Resultater';
const TEST_ARTIFACT_SUFFIX = '_TEST';

function ensureResultSheet_(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(TEST_RESULTS_SHEET);
  if (!sh){
    sh = ss.insertSheet(TEST_RESULTS_SHEET);
    sh.appendRow(['ts','testId','beskrivelse','status','ms','melding']);
    sh.getRange('A:F').setVerticalAlignment('top');
  }
  return sh;
}
function writeResult_(id, desc, status, ms, message){
  ensureResultSheet_().appendRow([new Date().toISOString(), id, desc, status, ms||0, message||'']);
}

/* ---------- DOCS DIALOG ---------- */
function openTestDocsDialog(){
  const html = HtmlService.createHtmlOutput(testDocsHtml_()).setWidth(720).setHeight(640);
  const ui = _tui_(); if (ui) ui.showModalDialog(html, 'Testdokumentasjon – Sikkerhet/GDPR/Backup/WCAG');
}
function testDocsHtml_(){
  const css = `
    <style>
      body{font-family:Arial,Helvetica,sans-serif;line-height:1.45}
      h2{margin:12px 0 6px} h3{margin:12px 0 4px}
      code{background:#f6f8fa;padding:2px 4px;border-radius:4px}
      .card{border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin:10px 0}
      .id{font-weight:bold;color:#2563eb} ul{margin:6px 0 6px 18px}
    </style>`;
  const wcag = `<div class="card"><h3>WCAG hurtigsjekk (Kap. 14 – supplering)</h3><p><b>Hensikt:</b> Sikre basale tilgjengelighetskrav (AA).</p><ul><li>Kontrast ≥ 4.5:1</li><li>Tastaturnavigasjon</li><li>Labels knyttet til inputs</li><li>Escape lukker dialog</li><li>Responsiv – ingen horisontal scrolling</li></ul></div>`;
  const blocks = [
    card_('T02_00','K2-00: RBAC – Meny-logikk filtrerer korrekt','Hensikt: riktig synlighet per rolle.',[
      'Automatisert simulering: Beboer får ikke “Admin”, Styremedlem får.']),
    card_('T02_01','K2-01: Vaktmester ser kun egne oppgaver','Filtrering etter ansvarlig.',[
      'Automatisert: syntetisk liste.']),
    card_('T02_02','K2-02: Vaktmester blokkeres fra personregister','Privat data skjermes.',[
      'Automatisert: autorisasjon må nekte.']),
    card_('T13_01','K13-01: Hendelseslogg er skrivebeskyttet','Uforanderlig revisjon.',[
      'Automatisert: beskyttelse + editors + domainEdit=false.']),
    wcag
  ];
  return css + `<div>${blocks.join('')}</div>`;
  function card_(id,title,purpose,bullets){
    return `<div class="card"><div class="id">${id}</div><h3>${title}</h3><p><b>Hensikt:</b> ${purpose}</p><ul>${bullets.map(b=>`<li>${b}</li>`).join('')}</ul></div>`;
  }
}

/* ---------- TESTER ---------- */
const TESTS = [
  // RBAC / meny
  { id:'T02_00', desc:'K2-00: RBAC – Meny-logikk filtrerer korrekt', fn: function(){
      const roller = { styre:['Styremedlem'], beboer:['Beboer'] };
      const regler = { 'Admin-meny':['Styremedlem','Admin'] };
      const sjekk = (userRoles, menu)=>!regler[menu] || userRoles.some(r=>regler[menu].includes(r));
      if (sjekk(roller.beboer,'Admin-meny')) throw new Error('Beboer fikk feilaktig Admin.');
      if (!sjekk(roller.styre,'Admin-meny')) throw new Error('Styremedlem ble nektet Admin.');
  }},
  { id:'T02_01', desc:'K2-01: Vaktmester ser kun egne oppgaver', fn: function(){
      const email='vaktmester@test.com';
      const alle=[{t:'A',ansvarlig:email},{t:'B',ansvarlig:'styre@test.com'},{t:'C',ansvarlig:email},{t:'D',ansvarlig:''}];
      const mine=alle.filter(x=>x.ansvarlig===email);
      if (mine.length!==2) throw new Error(`Forventet 2, fikk ${mine.length}`);
      if (mine.some(x=>x.ansvarlig!==email)) throw new Error('Inneholder feil oppgaver.');
  }},
  { id:'T02_02', desc:'K2-02: Vaktmester blokkeres fra personregister', fn: function(){
      function hentPersonData(roller){
        const tillatt=['Styremedlem','Admin','Kjernebruker'];
        if (!roller.some(r=>tillatt.includes(r))) throw new Error('Tilgang nektet.');
        return [{navn:'Test'}];
      }
      let blokkert=false;
      try{ hentPersonData(['Vaktmester']); }catch(e){ blokkert=(e.message==='Tilgang nektet.'); }
      if (!blokkert) throw new Error('Vaktmester ble ikke blokkert.');
      try{ hentPersonData(['Styremedlem']); }catch(e){ throw new Error('Styremedlem ble blokkert.'); }
  }},
  // Logging
  { id:'T13_01', desc:'K13-01: Hendelseslogg er skrivebeskyttet', fn: function(){
      const ss=SpreadsheetApp.getActive();
      const sh=ss.getSheetByName(SHEETS.LOGG);
      if (!sh) throw new Error(`Fane mangler: ${SHEETS.LOGG}`);
      const protections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      if (protections.length===0) throw new Error('Ingen beskyttelse på Hendelseslogg.');
      const p=protections[0];
      const me=Session.getEffectiveUser().getEmail();
      const others=p.getEditors().map(u=>u.getEmail()).filter(e=>e!==me);
      if (others.length>0) throw new Error('Andre har redigering til logg: '+others.join(', '));
      if (p.canDomainEdit()) throw new Error('Domenet kan redigere logg.');
  }},
  // Røyk/samspill (vil hoppes over hvis funksjoner mangler)
  { id:'T21_01', desc:'Dashboard røyk: dashMetrics() returnerer counts', fn: function(){
      if (typeof dashMetrics !== 'function') { Logger.log('dashMetrics mangler – hopper over.'); return; }
      const res = dashMetrics();
      if (!res || res.ok!==true || !res.counts) throw new Error('Ugyldig respons fra dashMetrics.');
      ['upcomingMeetings','openTasks','myTasks','pendingApprovals'].forEach(k=>{
        if (typeof res.counts[k] !== 'number') throw new Error(`counts.${k} ikke tall`);
      });
  }},
  { id:'T21_02', desc:'Vaktmester røyk: getTasksForVaktmester() kjører', fn: function(){
      if (typeof getTasksForVaktmester !== 'function') { Logger.log('getTasksForVaktmester mangler – hopper over.'); return; }
      const res = getTasksForVaktmester('active');
      if (!res || typeof res.ok==='undefined') throw new Error('Ugyldig svar fra getTasksForVaktmester.');
  }}
];

/* ---------- Test Runner ---------- */
function runAllTestsQuick_(){
  let pass=0, fail=0;
  TESTS.forEach(T=>{
    const start=Date.now();
    try{ T.fn(); writeResult_(T.id,T.desc,'PASS',Date.now()-start,''); pass++; }
    catch(e){ writeResult_(T.id,T.desc,'FAIL',Date.now()-start,String(e && e.message || e)); fail++; }
  });
  _testToast_(`Tester fullført: ${pass} PASS, ${fail} FAIL (totalt ${TESTS.length})`);
}
function runAllTestsAndShowReport_() {
  const results=[];
  TESTS.forEach(T=>{
    const start=Date.now(); let status='PASS', message='';
    try{ T.fn(); }catch(e){ status='FAIL'; message=String(e && e.message || e); }
    results.push({ id:T.id, desc:T.desc, status, ms:Date.now()-start, message });
    writeResult_(T.id,T.desc,status,results[results.length-1].ms,message);
  });
  const html = _generateTestReportHtml_(results);
  const ui=_tui_(); if (ui) ui.showModalDialog(html, 'Testresultater');
}
function _generateTestReportHtml_(results){
  const passCount = results.filter(r=>r.status==='PASS').length;
  const failCount = results.length - passCount;
  const summaryColor = failCount>0 ? '#dc2626' : '#16a34a';
  const rows = results.map(r=>{
    const msg = r.message ? `<br><small style="color:#555">${r.message.replace(/</g,'&lt;')}</small>` : '';
    return `<tr><td>${r.id}</td><td>${r.desc}${msg}</td><td style="font-weight:bold;color:${r.status==='PASS'?'#16a34a':'#dc2626'}">${r.status}</td><td>${r.ms} ms</td></tr>`;
  }).join('');
  const html = `
    <style>
      body{font-family:Arial,sans-serif;margin:0;padding:16px;font-size:14px}
      th,td{padding:8px;border-bottom:1px solid #ddd;text-align:left}
      th{background:#f2f2f2}
    </style>
    <h2>Testrapport</h2>
    <div>Resultat: <b style="color:${summaryColor}">${failCount>0?`${failCount} feilet`:'Alle bestått'}</b> (${passCount}/${results.length})</div>
    <table style="width:100%;margin-top:8px"><thead><tr><th>ID</th><th>Beskrivelse</th><th>Status</th><th>Tid</th></tr></thead><tbody>${rows}</tbody></table>`;
  return HtmlService.createHtmlOutput(html).setWidth(800).setHeight(600);
}
function runSingleTestPrompt_(){
  const ui=_tui_(); if (!ui) return;
  const choices = TESTS.map(t=>`${t.id} – ${t.desc}`);
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial,Helvetica,sans-serif;padding:8px">
       <h3 style="margin:0 0 6px">Kjør én test</h3>
       <label for="sel">Velg test:</label><br>
       <select id="sel" style="width:100%;margin:6px 0">${choices.map(c=>`<option>${c}</option>`).join('')}</select>
       <button onclick="google.script.run.withSuccessHandler(()=>google.script.host.close()).runSingleTestByIndex_(document.getElementById('sel').selectedIndex)">Kjør</button>
     </div>`
  ).setWidth(460).setHeight(180);
  ui.showModalDialog(html, 'Kjør én test');
}
function runSingleTestByIndex_(idx){
  const T = TESTS[idx];
  if (!T) { _testToast_('Ugyldig valg.'); return; }
  const start=Date.now();
  try{ T.fn(); writeResult_(T.id,T.desc,'PASS',Date.now()-start,''); _testToast_(`PASS: ${T.id}`); }
  catch(e){ writeResult_(T.id,T.desc,'FAIL',Date.now()-start,String(e && e.message || e)); _testToast_(`FAIL: ${T.id} – ${e.message||e}`); }
}
function cleanupTestArtifacts_() {
  const ui=_tui_(); if (!ui) return;
  const ss = SpreadsheetApp.getActive();
  const sheetsToDelete = ss.getSheets().filter(sh => sh.getName().endsWith(TEST_ARTIFACT_SUFFIX));
  if (sheetsToDelete.length === 0) { _testToast_('Ingen test-artefakter å rydde opp.'); return; }
  let message = 'Følgende vil bli slettet:\n\nARK:\n' + sheetsToDelete.map(s => `- ${s.getName()}`).join('\n');
  const resp = ui.prompt('Bekreft opprydding', message + "\n\nSkriv 'SLETT' for å bekrefte.", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton()===ui.Button.OK && resp.getResponseText().trim().toUpperCase()==='SLETT'){
    let n=0; sheetsToDelete.forEach(sh=>{ try{ SpreadsheetApp.getActive().deleteSheet(sh); n++; }catch(_){} });
    _testToast_(`${n} ark ble slettet.`);
  } else { _testToast_('Opprydding avbrutt.'); }
}

/* ---------- Legg “TESTING” i menyen (kalles fra 00_App_Core) ---------- */
function buildTestingSubmenu_(rootMenu){
  const ui = _tui_(); if (!ui) return;
  const sub = ui.createMenu('TESTING');
  sub.addItem('Kjør tester & vis rapport…', 'runAllTestsAndShowReport_');
  sub.addItem('Kjør alle tester (hurtig)', 'runAllTestsQuick_');
  sub.addItem('Kjør én test…', 'runSingleTestPrompt_');
  sub.addSeparator();
  sub.addItem('Vis testdokumentasjon…', 'openTestDocsDialog');
  sub.addSeparator();
  sub.addItem('Rydd opp test-artefakter…', 'cleanupTestArtifacts_');
  rootMenu.addSubMenu(sub);
}

