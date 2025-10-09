/**
 * FIL: 22_WebApp_Server.gs
 * FORMÅL: Backend logikk for dataaksess (henting/skriving) og forretningslogikk.
 * Leser fra spesifikt regneark via SPREADSHEET_ID.
 *
 * Ark som forventes (MÅ MATCHE NAVN I GOOGLE SHEET):
 * - "Brukere" (SHEET_USERS): Henter brukerrolle.
 * - "Meny" (SHEET_MENU): Henter menyvalg basert på rolle.
 * - "Avvik Kategorier" (SHEET_AVVIK_KATEGORIER): Henter kategorier for skjema.
 * - "Avvik Logg" (SHEET_AVVIK_LOGG): Logger innmeldte avvik.
 * - "Logg" (SHEET_LOG): Logger besøk/hendelser.
 */

const DATA = Object.freeze({
    // !!! VIKTIG: ERSTATT MED DIN EKTE SPREADSHEET ID !!!
    SPREADSHEET_ID: '1v91oJ7F5qiZHbp6trHnRiC5kKyWYXPklvIsVvC2lVao', // <- Sjekk denne!
    SHEET_USERS: 'Brukere',
    SHEET_MENU:  'Meny',
    SHEET_LOG:   'Logg',
    SHEET_AVVIK_KATEGORIER: 'Avvik Kategorier',
    SHEET_AVVIK_LOGG: 'Avvik Logg',
});

// ------------------------------------------------------------------
// Hjelpere for datatilgang
// ------------------------------------------------------------------

/** Åpner regnearket. Kritiske feil stoppes her. */
function _ss_() {
    try {
        return SpreadsheetApp.openById(DATA.SPREADSHEET_ID);
    } catch (e) {
        // Kritisk feil ved tilgang/ID
        Logger.log('KRITISK FEIL: Klarte ikke åpne regneark med ID: %s. Feil: %s', DATA.SPREADSHEET_ID, e.message);
        throw new Error('Klarte ikke koble til databasen. Sjekk SPREADSHEET_ID og tilgang.');
    }
}

/** Henter et ark trygt, eller logger feil. */
function _getSheetByName_(name) {
    const ss = _ss_();
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
        Logger.log('ADVARSEL: Mangler ark "%s" i regnearket.', name);
        // Tillater flyt å fortsette ved å returnere null, som må håndteres av kallende funksjon
    }
    return sheet;
}

/** Trygg uthenting av brukerens e-post. */
function _getUserEmail_() {
    try {
        const a = Session.getActiveUser() && Session.getActiveUser().getEmail();
        const b = Session.getEffectiveUser && Session.getEffectiveUser().getEmail();
        return String(a || b || '').trim();
    } catch (e) {
        return '';
    }
}

/** Case-insensitive headeroppslag med støtte for flere navn. */
function _indexByNames_(header, wanted) {
    const lower = header.map(function(h) { return String(h).toLowerCase().trim(); });
    const out = {};
    Object.keys(wanted).forEach(function(key) {
        const candidates = wanted[key].map(function(x) { return String(x).toLowerCase().trim(); });
        let idx = -1;
        for (let i = 0; i < lower.length; i++) {
            if (candidates.indexOf(lower[i]) !== -1) { 
                idx = i; 
                break; 
            }
        }
        out[key] = idx;
    });
    return out;
}

function _parseRoles_(cell) {
    if (!cell) return ['Gjest'];
    if (Array.isArray(cell)) cell = cell.join(',');
    return String(cell).split(/[,;]/).map(function(s) { return s.trim(); }).filter(Boolean);
}

// ------------------------------------------------------------------
// Hoved-API for frontend
// ------------------------------------------------------------------

/** Hentes ved app-oppstart for å autentisere og bygge meny. */
function uiBootstrap() {
    Logger.log('[uiBootstrap] Starter opplasting...');
    try {
        const email = _getUserEmail_();
        if (!email) {
            // Hvis e-post er tom (f.eks. ved anonym tilgang eller feil), faller vi tilbake til Gjest.
             Logger.log('[uiBootstrap] Finner ikke aktiv bruker-e-post. Antar Gjest.');
        }

        const user  = _getUserRecord_(email);
        const menu  = _getMenuForRoles_(user.roles);
        
        // Logg kun om vi har en e-post.
        if (email) {
            logUserVisit_(user.email, user.name, user.roles);
        }

        Logger.log('[uiBootstrap] Ferdig. Bruker: %s, Roller: %s', user.name, user.roles.join(', '));
        return { user: user, menu: menu };

    } catch (err) {
        Logger.log('KRITISK uiBootstrap feil: %s', err.message, err.stack);
        const guest = { name: 'Gjest', email: 'ukjent', roles: ['Gjest'] };
        // Returnerer minimumsdata for at klientsiden kan vise feilmelding
        return { 
            user: guest, 
            menu: [{ name: 'Feil', action: 'dashboard' }],
            error: err.message
        };
    }
}

/** Leser "Brukere" (Epost | Navn | Roller) og returnerer {name,email,roles[]}. */
function _getUserRecord_(email) {
    const sh = _getSheetByName_(DATA.SHEET_USERS);
    if (!sh || sh.getLastRow() < 2) {
        Logger.log('[Brukere] ark mangler eller tomt. Standard Gjest.');
        return { name: email || 'Gjest', email: email || '', roles: ['Gjest'] };
    }

    const emailNorm = String(email || '').trim().toLowerCase().replace(/\s+/g, '');
    if (!emailNorm) {
        return { name: 'Gjest', email: '', roles: ['Gjest'] };
    }

    const values = sh.getDataRange().getValues();
    const header = values.shift().map(String);
    const idx = _indexByNames_(header, {
        Epost:  ['Epost','Email','E-mail','E post','E-post'],
        Navn:   ['Navn','Name'],
        Roller: ['Roller','Roles']
    });

    const row = values.find(function(r) { return String(r[idx.Epost] || '').trim().toLowerCase().replace(/\s+/g, '') === emailNorm; });
    if (!row) {
        Logger.log('Bruker %s ikke funnet i listen. Tilbyr Gjest-rolle.', email);
        return { name: email || 'Gjest', email: email || '', roles: ['Gjest'] };
    }

    const name  = row[idx.Navn] || email || 'Ukjent';
    const roles = _parseRoles_(row[idx.Roller]);
    return { name: String(name), email: String(email), roles: roles };
}

/** Leser "Meny"-arket og filtrerer på brukerroller. */
function _getMenuForRoles_(roles) {
    roles = Array.isArray(roles) ? roles.slice() : ['Gjest'];

    // Arving av roller (Admin > Styremedlem > Beboer/Vaktmester)
    if (roles.includes('Admin')) {
        roles = roles.concat(['Styremedlem', 'Beboer', 'Vaktmester']).filter(function(r, i, arr) { return arr.indexOf(r) === i; });
    }
    if (roles.includes('Styremedlem')) {
        roles = roles.concat(['Beboer']).filter(function(r, i, arr) { return arr.indexOf(r) === i; });
    }

    const sh = _getSheetByName_(DATA.SHEET_MENU);
    if (!sh) return [];

    const values = sh.getDataRange().getValues();
    const header = values.shift().map(String);
    const idx = _indexByNames_(header, {
        Rolle:    ['Rolle','Roller','Role'],
        Meny:     ['Meny','Menu','Navn','Name'],
        Funksjon: ['Funksjon','Function','Action'],
        Sort:     ['Sort','Order','Idx']
    });

    let rows = values.filter(function(r) {
        const role = String(r[idx.Rolle] || '').trim().toLowerCase();
        return roles.some(function(userRole) { return role === String(userRole).trim().toLowerCase(); });
    });

    // Sortering (i minnet)
    if (idx.Sort !== -1) {
        rows = rows.sort(function(a, b) { return (Number(a[idx.Sort] || 0)) - (Number(b[idx.Sort] || 0)); });
    }

    // Fjern duplikater og formatter til objekt
    const seen = new Set();
    const out = [];
    rows.forEach(function(r) {
        const name   = String(r[idx.Meny] || '').trim();
        const action = String(r[idx.Funksjon] || '').trim();
        if (!name || !action) return;
        const key = name + '|' + action;
        if (!seen.has(key)) { 
            seen.add(key); 
            out.push({ name: name, action: action }); 
        }
    });

    return out;
}

/** Henter kategorier for avviksmelding fra arket. */
function getAvvikCategories() {
    Logger.log('[getAvvikCategories] Henter kategorier...');
    const sh = _getSheetByName_(DATA.SHEET_AVVIK_KATEGORIER);
    if (!sh) {
        throw new Error('Mangler arket "Avvik Kategorier". Sjekk navnet.');
    }

    try {
        const values = sh.getDataRange().getValues();
        const header = values.shift().map(String);
        
        // Definer forventede kolonnenavn
        const idx = _indexByNames_(header, {
            ID: ['KategoriID', 'ID'],
            Navn: ['Navn', 'Name'],
            Prioritet: ['Prioritet', 'Prio', 'Priority']
        });
        
        if (idx.ID === -1 || idx.Navn === -1 || idx.Prioritet === -1) {
             throw new Error('Mangler en eller flere kritiske kolonner (KategoriID, Navn, Prioritet) i "Avvik Kategorier".');
        }

        const categories = values.map(function(row) {
            return {
                id: String(row[idx.ID] || ''),
                name: String(row[idx.Navn] || ''),
                priority: String(row[idx.Prioritet] || '')
            };
        }).filter(function(cat) { return cat.id && cat.name; }); // Filtrer bort tomme rader
        
        Logger.log('[getAvvikCategories] Fant %d kategorier.', categories.length);
        return categories;
        
    } catch (e) {
        Logger.log('Feil ved lesing av avvikskategorier: %s', e.message);
        throw new Error('Klarte ikke hente avvikskategorier: ' + e.message);
    }
}

/** Tar imot data fra klientsiden for å registrere et nytt avvik i Sheets. */
function createAvvik(data) {
    data = data || {};
    Logger.log('[createAvvik] Mottatt data: %s', JSON.stringify(data));
    
    // FIKS: Bruker data-objektet direkte istedenfor å parse data.args.
    const form = data; 
    const userEmail = _getUserEmail_() || 'gjest@ukjent.no';

    const sh = _getSheetByName_(DATA.SHEET_AVVIK_LOGG);
    if (!sh) {
        throw new Error('Mangler arket "Avvik Logg". Kan ikke registrere avvik.');
    }

    // Unik ID genereres i backend
    const avvikId = Utilities.getUuid().substring(0, 8).toUpperCase(); 

    try {
        // Legg til en rad med data
        const newRow = [
            new Date(),
            userEmail,
            form.category || 'Ukjent',
            form.location || 'Ukjent sted',
            form.description || 'Ingen beskrivelse gitt.',
            'Ny', // Standard status
            avvikId // Unik ID
        ];
        
        sh.appendRow(newRow);
        Logger.log('[createAvvik] Nytt avvik registrert: %s', avvikId);
        
        return { success: true, avvikId: avvikId };

    } catch (e) {
        Logger.log('Feil ved skriving til Avvik Logg: %s', e.message);
        throw new Error('Klarte ikke lagre avvik: ' + e.message);
    }
}

/** Logger besøk til "Logg". */
function logUserVisit_(email, name, roles) {
  try {
    const sh = _getSheetByName_(DATA.SHEET_LOG);
    if (sh) {
        sh.appendRow([new Date(), String(email||''), String(name||''), (roles||[]).join(', ')]);
    }
  } catch (e) {
    Logger.log('ADVARSEL: Feil ved loggføring av besøk: %s', e.message);
    // Stille feil: lar appen fortsette selv om logger feiler
  }
}

// ------------------------------------------------------------------
// Debug
// ------------------------------------------------------------------
function debugBootstrap() {
  const result = uiBootstrap();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * En superenkel funksjon for å teste den grunnleggende koblingen til regnearket.
 */
function debugSheetConnection() {
  try {
    Logger.log("DEBUG: Starter test...");
    // Pass på at denne ID-en er 100% korrekt!
    const ss = SpreadsheetApp.openById("1v91oJ7F5qiZHbp6trHnRiC5kKyWYXPklvIsVvC2lVao"); 
    Logger.log("DEBUG: Fikk åpnet regneark OK.");
    
    const sheet = ss.getSheetByName("Brukere");
    if (!sheet) {
      Logger.log("DEBUG: FANT IKKE ARKET 'Brukere'");
      return "FEIL: Fant ikke arket 'Brukere'. Sjekk at navnet er helt likt.";
    }
    Logger.log("DEBUG: Fikk tak i arket 'Brukere' OK.");

    const cellValue = sheet.getRange("A1").getValue();
    Logger.log("DEBUG: Leste verdi fra A1: " + cellValue);
    return "SUCCESS: Leste verdien '" + cellValue + "' fra celle A1 i arket 'Brukere'. Koblingen fungerer!";
    
  } catch (e) {
    Logger.log("DEBUG: KRITISK FEIL I TESTFUNKSJON: " + e.toString());
    return "KRITISK FEIL: " + e.message;
  }
}