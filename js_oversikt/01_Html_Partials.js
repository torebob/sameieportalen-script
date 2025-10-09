function include(filename, data) {
  var t = HtmlService.createTemplateFromFile(filename);
  if (data && typeof data === 'object') {
    for (var k in data) if (Object.prototype.hasOwnProperty.call(data, k)) t[k] = data[k];
  }
  return t.evaluate().getContent();
}
/**
 * Midlertidig stub for å teste frontend.
 * Returnerer falske brukerdata til frontenden.
 */
function uiBootstrap() {
  return {
    user: {
      name: "Testbruker",
      email: "test@example.com",
      roles: ["Styremedlem", "Admin"]
    }
  };
}
