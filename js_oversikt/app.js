<!-- public/js/app.html -->
<script>
// Globalt objekt for frontend-tilstand
const AppState = {
    user: { roles: ['Gjest'], name: 'Laster...' },
    menu: [],
    currentPage: null,
    // Registrerte sider (action-navn: filnavn i public/js/ui/mappen)
    // OBS: Disse navnene må matche HTML-filene (f.eks. 'styret' matcher public/js/ui/styret-page.html)
    pages: {
        'dashboard': 'public/js/ui/dashboard-page',
        'styret': 'public/js/ui/styret-page',
        'booking': 'public/js/ui/booking-page',
        'dokumenter': 'public/js/ui/dokumenter-page',
        'avvik-new': 'public/js/ui/avvik-new-page',
        // Legg til flere sider her etter hvert
    },
    pageContent: {} // Cache for innhold som lastes asynkront
};

/**
 * Hjelpefunksjon for å kalle Apps Script backend asynkront.
 * Dette er kritisk for all kommunikasjon med Sheets/Google Services.
 * @param {string} functionName Navnet på Apps Script-funksjonen (i .gs-filen)
 * @param {object} args Eventuelle argumenter til funksjonen
 * @returns {Promise<any>}
 */
function callServer(functionName, args = {}) {
    const spinner = document.getElementById('app-spinner');
    spinner.style.display = 'flex';

    // Returner en promise for å håndtere asynkrone kall
    return new Promise(function(resolve, reject) { // Bruker standard function her

        // Definer suksess-handler
        const successHandler = function(response) {
            spinner.style.display = 'none';

            if (response && response.success === true) {
                return resolve(response.result);
            } else if (response && response.error) {
                console.error('Backend Error:', response.error);
                // Bruk en custom modal/toast i prod.
                // alert('Feil fra server: ' + response.error);
                return reject(new Error(response.error));
            }
            return reject(new Error('Ukjent feil ved serverkall.'));
        }

        // Definer feil-handler
        const failureHandler = function(error) {
            spinner.style.display = 'none';
            console.error('Server call failed:', error);
            // Bruk en custom modal/toast i prod.
            // alert('En uventet feil oppstod: ' + error);
            reject(error);
        };

        // Kaller Apps Script-funksjonen, doPost er ruteren på serversiden.
        google.script.run
        .withSuccessHandler(successHandler)
        .withFailureHandler(failureHandler)
        .doPost({ function: functionName, ...args });
    });
}

/**
 * Laster inn en spesifikk side i hovedinnholdsområdet.
 * @param {string} action Navn på aksjonen/siden (f.eks. 'styret')
 */
function loadPage(action) {
    const contentArea = document.getElementById('content-area');
    const pagePath = AppState.pages[action];
    const pageName = action.charAt(0).toUpperCase() + action.slice(1);

    if (!pagePath) {
        contentArea.innerHTML = `<h2 class="text-xl font-bold text-red-600 p-6">Feil: Siden '${action}' eksisterer ikke eller er ikke definert i AppState.pages.</h2>`;
        AppState.currentPage = null;
        return;
    }

    AppState.currentPage = action;
    contentArea.innerHTML = `<div class="p-6"><h2 class="text-2xl font-semibold text-gray-700">Laster: ${pageName}...</h2></div>`;

    // 1. Hent HTML-innholdet for siden via getHtmlContent
    callServer('getHtmlContent', { filename: pagePath })
    .then(function(result) { // Bruker standard function her
        // result.content inneholder HTML-stringen
        contentArea.innerHTML = result.content || `<div class="p-6">Kunne ikke laste innhold for ${pageName}.</div>`;
        document.title = 'Sameieportalen - ' + pageName;
        // Skroller til toppen av innholdsområdet etter lasting
        contentArea.scrollTop = 0;
    })
    .catch(function(error) { // Bruker standard function her
        contentArea.innerHTML = `<div class="p-6"><h2 class="text-xl font-bold text-red-600">Feil ved lasting av side: ${error.message}</h2><p>Sjekk Apps Script Logger for detaljer.</p></div>`;
    });

    // Oppdater hash i URL-en
    window.location.hash = '#' + action;
    updateUI();
}

/** Oppdaterer topplinjen med brukerdata og genererer meny. */
function updateUI() {
    const userInfoEl = document.getElementById('user-info');
    const menuEl = document.querySelector('#main-menu ul');

    userInfoEl.textContent = `${AppState.user.name} (${AppState.user.roles.join(', ')})`;

    // Bygg menyen basert på roller fra backend
    menuEl.innerHTML = '';
    AppState.menu.forEach(function(item) { // Bruker standard function her
        const li = document.createElement('li');
        // Legger til dynamisk ruting for lenker i sidemenyen
        li.innerHTML = `<a href="#${item.action}" class="block px-4 py-2 text-sm text-gray-300 hover:bg-blue-700 transition duration-150 rounded-lg">${item.name}</a>`;
        li.querySelector('a').addEventListener('click', function(e) { // Bruker standard function her
            e.preventDefault();
            loadPage(item.action);
        });
        menuEl.appendChild(li);
    });

    // Legg til hendelseslytter for hash-endringer for å støtte tilbake/frem-knapper
    window.addEventListener('hashchange', function() { // Bruker standard function her
        const hash = window.location.hash.substring(1);
        if (hash && hash !== AppState.currentPage) {
            loadPage(hash);
        }
    });

    // Legg til hendelseslytter for lenker i content-area som bruker #hash-ruting
    document.getElementById('content-area').addEventListener('click', function(e) { // Bruker standard function her
        let target = e.target;
        // Gå oppover i DOM-en hvis klikket skjedde på et barne-element av <a>
        while (target && target.tagName !== 'A') {
            target = target.parentElement;
        }

        if (target && target.href && target.href.includes('#')) {
            const hash = target.hash.substring(1);
            if (AppState.pages[hash]) {
                e.preventDefault();
                loadPage(hash);
            }
        }
    });
}

/** Initialiserer appen ved start. */
function initApp() { // Fjerner async for å være mer kompatibel, selv om Apps Script's google.script.run er asynkron
    console.log("App starter...");
    const contentArea = document.getElementById('content-area');

    // Gjøres asynkront via callServer
    callServer('uiBootstrap')
    .then(function(data) { // Bruker standard function her
        AppState.user = data.user;
        AppState.menu = data.menu;

        updateUI();

        const initialAction = window.location.hash ? window.location.hash.substring(1) : AppState.menu[0]?.action || 'dashboard';
        const finalAction = AppState.pages[initialAction] ? initialAction : 'dashboard';

        loadPage(finalAction);

    })
    .catch(function(e) { // Bruker standard function her
        console.error("Kritisk oppstartsfeil i initApp:", e);
        const errorMsg = 'Kunne ikke koble til serveren. Sjekk Apps Script Logger og publisering.';
        document.getElementById('app-spinner').style.display = 'none';
        contentArea.innerHTML = `<div class="p-6"><h2 class="text-xl font-bold text-red-600">${errorMsg}</h2><p class="text-gray-600">${e.message}</p></div>`;
    });

    // Skjules inne i promise-handleren, men vi sikrer at den ikke spinner i evig tid ved catch
}

// Kjør initialiseringsfunksjonen
window.onload = initApp;

</script>
<!-- Simuler Google Script API for lokal testing/intelliSense -->
<script>
// Denne seksjonen er kun for utvikling/testing utenfor Apps Script-miljøet
if (typeof google === 'undefined' || typeof google.script === 'undefined') {
    window.google = {
        script: {
            run: {
                withSuccessHandler: function(handler) { // Bruker standard function her
                    return {
                        withFailureHandler: function(fail) { // Bruker standard function her
                            return {
                                doPost: function(args) { // Bruker standard function her
                                    // MOCK-implementasjon for doPost
                                    setTimeout(function() { // Bruker standard function her
                                        if (args.function === 'uiBootstrap') {
                                            const mockData = {
                                                user: { name: 'Mock Bruker', email: 'test@sameie.no', roles: ['Styreleder', 'Beboer'] },
                                                menu: [
                                                    { name: 'Dashboard', action: 'dashboard' },
                                                    { name: 'Styret', action: 'styret' },
                                                    { name: 'Booking', action: 'booking' },
                                                    { name: 'Dokumenter', action: 'dokumenter' }
                                                ]
                                            };
                                            handler({ success: true, result: mockData });
                                        } else if (args.function === 'getHtmlContent') {
                                            const pageKey = args.filename;
                                            const contentMap = {
                                                'public/js/ui/dashboard-page': '<!-- MOCK --> <h2 class="text-3xl font-light">MOCK Dashboard</h2><p>Dette er mock data.</p>',
                                                'public/js/ui/styret-page': '<!-- MOCK --> <h2 class="text-3xl font-light">MOCK Styret</h2><p>Styret side.</p>',
                                                'public/js/ui/avvik-new-page': '<!-- MOCK --> <h2 class="text-3xl font-light">MOCK Nytt Avvik</h2><p>Melde skjema.</p>',
                                                'public/js/ui/booking-page': '<!-- MOCK --> <h2 class="text-3xl font-light">MOCK Booking</h2><p>Booking side.</p>',
                                                'public/js/ui/dokumenter-page': '<!-- MOCK --> <h2 class="text-3xl font-light">MOCK Dokumenter</h2><p>Dokumenter side.</p>',
                                            };
                                            const htmlContent = contentMap[pageKey] || `<h2 class="text-3xl font-light">MOCK: Siden for ${pageKey} finnes ikke.</h2>`;

                                            // Må etterligne Apps Script-responsen:
                                            const response = { success: true, result: { content: htmlContent } };

                                            handler(response);
                                        } else {
                                            fail('MOCK: Ukjent funksjon: ' + args.function);
                                        }
                                    }, 100);
                                }
                            };
                        }
                    };
                }
            }
        }
    };
    // Sørg for at initApp kalles hvis vi er i en mock-miljø
    if (document.readyState === 'complete') {
        initApp();
    } else {
        window.addEventListener('load', initApp);
    }
}
</script>
