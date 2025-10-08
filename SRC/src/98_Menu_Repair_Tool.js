/* Hurtigverktøy for meny-problemer */

function forceShowMenu(){
  // Kjør denne én gang fra editoren hvis menyen ikke vises
  spOnOpen({});
  Logger.log('Meny tvunget frem med spOnOpen()');
}

function repairMenuTrigger(){
  // Installer en "installable" onOpen-trigger som kaller spOnOpen uansett
  var ss = SpreadsheetApp.getActive();
  var ssId = ss.getId();

  // Fjern gamle spOnOpen/onOpen-triggere for dette arket
  ScriptApp.getProjectTriggers().forEach(function(t){
    var h = t.getHandlerFunction && t.getHandlerFunction();
    var src = t.getTriggerSourceId && t.getTriggerSourceId();
    if ((h === 'spOnOpen' || h === 'onOpen') && src === ssId) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Ny trigger
  ScriptApp.newTrigger('spOnOpen').forSpreadsheet(ssId).onOpen().create();
  Logger.log('Installable onOpen-trigger opprettet for spOnOpen()');
}

function listProjectTriggers(){
  return ScriptApp.getProjectTriggers().map(function(t){
    return {
      handler: t.getHandlerFunction && t.getHandlerFunction(),
      type: t.getEventType && String(t.getEventType()),
      sourceId: t.getTriggerSourceId && t.getTriggerSourceId()
    };
  });
}
