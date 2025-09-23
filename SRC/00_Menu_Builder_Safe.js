
/* ====================== Menybygger (robust) ====================== */
/* FILE: 00_Menu_Builder_Safe.gs | VERSION: 1.3.0 | UPDATED: 2025-09-14 */

function onOpen(){ spOnOpen(); }

function spOnOpen(){
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Sameieportalen');

  menu.addItem('Møteoversikt & Protokoller', 'openMeetingsUI');

  // Vaktmester (viser bare hvis filen finnes)
  try{
    if (typeof openVaktmesterUI === 'function'){
      menu.addItem('Mine Oppgaver (Vaktmester)', 'openVaktmesterUI');
    }
  }catch(_){/* ignore */}

  menu.addSeparator();
  menu.addItem('Opprett basisfaner', 'createBaseSheets');
  menu.addItem('Åpne Dashboard', 'openDashboard');
  menu.addItem('Kjør kvalitetssjekk', 'runAllChecks');

  // Admin undermeny (vises alltid, funksjoner kan mangle – det tåler vi)
  const admin = ui.createMenu('Admin');
  admin.addItem('Tøm dashboard-cache', 'clearDashboardCache');
  admin.addItem('Force: vis meny nå', 'forceShowMenu');
  menu.addSubMenu(admin);

  menu.addToUi();
  Logger.log('Meny tvunget frem med spOnOpen()');
}

// Nødhjelp dersom menyen er borte i filen
function forceShowMenu(){ spOnOpen(); }
