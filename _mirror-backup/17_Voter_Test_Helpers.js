/* ====================== Voter – Test Helpers (uten konsoll) ======================
 * FILE: 17_Voter_Test_Helpers.gs | VERSION: 1.0.0 | UPDATED: 2025-09-14
 * FORMÅL: Kjør stemmemodulen via "Run/Kjør" og en engangsmeny – ingen konsoll nødvendig.
 * Gir meny: Voter (dev) med kommandoer for hurtigtest.
 * Avhenger av 16_Voter_Modul.gs.
 * ================================================================================ */

function voterDevSetupSheets(){
  const res = voterEnsureSheets_();
  SpreadsheetApp.getUi().alert('OK', 'Stemmer/SAK-ark er klar:\n' + JSON.stringify(res), SpreadsheetApp.getUi().ButtonSet.OK);
}

function voterDevSaveVote(){
  const ui = SpreadsheetApp.getUi();
  const id = ui.prompt('Møte-ID', 'F.eks. M-2025-10-01', ui.ButtonSet.OK_CANCEL);
  if (id.getSelectedButton() !== ui.Button.OK) return;
  const s = ui.prompt('Saksnr', 'F.eks. S-0012025', ui.ButtonSet.OK_CANCEL);
  if (s.getSelectedButton() !== ui.Button.OK) return;
  const v = ui.prompt('Stemme', 'Skriv JA, NEI eller BLANK', ui.ButtonSet.OK_CANCEL);
  if (v.getSelectedButton() !== ui.Button.OK) return;

  try{
    const r = voterSaveVote(id.getResponseText().trim(), s.getResponseText().trim(), v.getResponseText().trim());
    ui.alert('Lagret', r.message || 'OK', ui.ButtonSet.OK);
  }catch(e){
    ui.alert('Feil', e.message, ui.ButtonSet.OK);
  }
}

function voterDevShowStatus(){
  const ui = SpreadsheetApp.getUi();
  const id = ui.prompt('Møte-ID', 'F.eks. M-2025-10-01', ui.ButtonSet.OK_CANCEL);
  if (id.getSelectedButton() !== ui.Button.OK) return;
  const s = ui.prompt('Saksnr', 'F.eks. S-0012025', ui.ButtonSet.OK_CANCEL);
  if (s.getSelectedButton() !== ui.Button.OK) return;

  try{
    const r = voterGetStatus(id.getResponseText().trim(), s.getResponseText().trim());
    const msg =
      `Møte: ${r.moteId}\nSak: ${r.saksnr}\n\n` +
      `Din stemme: ${r.myVote||'—'}\n` +
      `Tellinger: JA=${r.counts?.JA||0}, NEI=${r.counts?.NEI||0}, BLANK=${r.counts?.BLANK||0}\n` +
      `Låst: ${r.locked ? 'Ja' : 'Nei'}\n` +
      (r.vedtak ? `Vedtak:\n${r.vedtak}` : '');
    ui.alert('Status', msg, ui.ButtonSet.OK);
  }catch(e){
    ui.alert('Feil', e.message, ui.ButtonSet.OK);
  }
}

function voterDevLockDecision(){
  const ui = SpreadsheetApp.getUi();
  const id = ui.prompt('Møte-ID', 'F.eks. M-2025-10-01', ui.ButtonSet.OK_CANCEL);
  if (id.getSelectedButton() !== ui.Button.OK) return;
  const s = ui.prompt('Saksnr', 'F.eks. S-0012025', ui.ButtonSet.OK_CANCEL);
  if (s.getSelectedButton() !== ui.Button.OK) return;
  const t = ui.prompt('Vedtakstekst (låses)', 'Skriv endelig vedtakstekst slik den skal stå i protokollen.', ui.ButtonSet.OK_CANCEL);
  if (t.getSelectedButton() !== ui.Button.OK) return;

  try{
    const r = voterLockDecision(id.getResponseText().trim(), s.getResponseText().trim(), t.getResponseText().trim());
    ui.alert('Vedtak låst', r.message || 'OK', ui.ButtonSet.OK);
  }catch(e){
    ui.alert('Feil', e.message, ui.ButtonSet.OK);
  }
}

/** Engangsmeny – kjør denne fra Run/Kjør for å få menyen “Voter (dev)”. */
function voterDevShowMenuOnce(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Voter (dev)')
    .addItem('Opprett stemme-ark', 'voterDevSetupSheets')
    .addItem('Stem JA/NEI/BLANK…', 'voterDevSaveVote')
    .addItem('Vis status…', 'voterDevShowStatus')
    .addItem('Lås vedtak…', 'voterDevLockDecision')
    .addToUi();
  ui.alert('Meny lagt til', 'Se menylinjen: “Voter (dev)”.', ui.ButtonSet.OK);
}
