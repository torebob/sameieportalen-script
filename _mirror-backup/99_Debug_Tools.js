/* ====================== Debug / Admin Tools (nødfunksjon) ====================== */
/* FILE: 99_Debug_Tools.gs | VERSION: 1.1.0 | UPDATED: 2025-09-13 */
/* FORMÅL: Admin-only “nødluke” for å åpne ett ark midlertidig. Arket skjules
   automatisk ved neste onOpen. Ingen “vis/skjul utviklerfaner”-brytere. */

/* --- List alle ark med status --- */
function devListSheets(){
  const ss = SpreadsheetApp.getActive();
  const list = ss.getSheets().map((sh,i)=>({
    idx: i+1, name: sh.getName(), hidden: sh.isSheetHidden(), id: sh.getSheetId()
  }));
  Logger.log('Ark funnet:\n' + list.map(r =>
    `${r.idx}: ${r.name}${r.hidden?' (SKJULT)':''}`
  ).join('\n'));
  return list;
}

/* --- Admin-only: dialog for å velge ark å åpne midlertidig --- */
function adminOpenSheetTemporarilyDialog(){
  try{
    if (typeof __core_isAdmin === 'function' ? !__core_isAdmin() : true){
      _alert_('Du må være i ADMIN_WHITELIST (Konfig) for å bruke denne funksjonen.', 'Ingen tilgang');
      return;
    }
    const ss = SpreadsheetApp.getActive();
    const sheets = ss.getSheets();
    // List kandidater = alle tekniske + de som er skjult
    const tech = (typeof TECHNICAL_SHEETS !== 'undefined') ? TECHNICAL_SHEETS : [];
    const candidates = sheets.filter(sh => sh.isSheetHidden() || tech.includes(sh.getName()));

    const optionsHtml = candidates.map(sh =>
      `<option value="${String(sh.getSheetId())}">${_htmlEsc_(sh.getName())}</option>`
    ).join('');

    const html = HtmlService.createHtmlOutput(
      `<!DOCTYPE html><html><head><meta charset="utf-8">
        <style>
          body{font-family:Arial,Helvetica,sans-serif;margin:16px;}
          label{display:block;margin-bottom:6px;}
          select,button{font-size:14px;padding:6px 8px;}
          .row{margin-top:12px;}
        </style>
      </head><body>
        <h3>Åpne ark midlertidig (admin)</h3>
        <label for="sheet">Velg ark som skal åpnes:</label>
        <select id="sheet">${optionsHtml}</select>
        <div class="row">
          <button onclick="openNow()">Åpne</button>
          <button onclick="google.script.host.close()">Avbryt</button>
        </div>
        <p style="color:#64748b;margin-top:12px">
          Arket blir automatisk skjult igjen ved neste åpning av filen.
        </p>
        <script>
          function openNow(){
            var id = document.getElementById('sheet').value;
            google.script.run
              .withSuccessHandler(function(msg){ alert(msg); google.script.host.close(); })
              .withFailureHandler(function(err){ alert('Feil: '+err.message); })
              .adminOpenSheetTemporarily_(id);
          }
        </script>
      </body></html>`
    ).setWidth(420).setHeight(240);
    _ui().showModalDialog(html, 'Åpne ark (midlertidig)');
  }catch(err){
    _alert_('Klarte ikke å åpne dialog: '+err, 'Feil');
  }
}

/* --- Server-side: åpne og markere for re-skjul ved neste onOpen --- */
function adminOpenSheetTemporarily_(sheetId){
  if (typeof __core_isAdmin === 'function' ? !__core_isAdmin() : true){
    throw new Error('Ingen tilgang (ADMIN_WHITELIST).');
  }
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheets().find(s => String(s.getSheetId()) === String(sheetId));
  if (!sh) throw new Error('Fant ikke arket.');

  try{
    sh.showSheet();
    ss.setActiveSheet(sh);
    PROPS.setProperty(PROP_KEYS.TEMP_VIS_SHEET_ID, String(sh.getSheetId()));
    return `Ark "${sh.getName()}" er åpnet midlertidig og vil skjules ved neste åpning.`;
  }catch(err){
    throw new Error('Kunne ikke åpne arket: '+err);
  }
}

/* --- Små hjelpefunksjoner --- */
function _htmlEsc_(s){
  return String(s||'')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

/* (valgfritt) Rydd opp manuelt hvis ønskelig */
function adminHideTempSheetNow(){
  if (typeof __core_rehideTemp === 'function') __core_rehideTemp();
  _alert_('Evt. midlertidig ark er forsøkt skjult nå.');
}
