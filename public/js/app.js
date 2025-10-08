<script>
// =================================================================
// Sameieportalen Frontend - Hovedapplikasjon (Hjernen)
// FILE: public/js/app.js (limt inn i public_js_app.html)
// =================================================================

// ---------- DOM-elementer ----------
const contentArea = document.getElementById('content-area');
const mainMenuUl  = document.querySelector('#main-menu ul');
const userInfoDiv = document.getElementById('user-info');

// ---------- Meny-konfig (fallback basert på ROLLER) ----------
// Brukes KUN hvis backend (uiBootstrap) ikke sender en ferdig meny.
// action peker på *funksjonsnavn* vi har i window (se lenger ned).
const MENU_CONFIG = [
    { name: 'Dashboard',           roles: ['Gjest','Beboer','Seksjonseier','Leietaker','Styremedlem','Admin','Vaktmester'], action: 'loadDashboard' },
{ name: 'Oppslag',             roles: ['Gjest','Beboer','Seksjonseier','Leietaker','Styremedlem','Admin','Vaktmester'], action: 'loadAnnouncements' },
// ------------------ STYRET ------------------
{ name: 'Møter & Protokoller', roles: ['Styremedlem','Admin'], action: 'loadMeetings' },
{ name: 'Send Oppslag',        roles: ['Styremedlem','Admin'], action: 'openNewAnnouncementUI' },
{ name: 'Brukeradmin',         roles: ['Admin'],               action: 'loadUserAdmin' },
// ------------------ VAKTMESTER ------------------
{ name: 'Mine Oppgaver',       roles: ['Vaktmester'],          action: 'loadJanitorTasks' },
];

// ---------- App-Initialisering ----------
function init() {
    setLoadingState('Laster brukerinformasjon…');

    google.script.run
    .withSuccessHandler(userData => {
        // Forventet format: { user: {name,email,roles:[]}, menu?: [{name,action}] }
        const user = userData && userData.user ? userData.user : { name:'Gjest', email:'', roles:['Gjest'] };
        renderUserInfo(user);

        // Dynamisk meny: bruk server-meny hvis den finnes, ellers bygg fra roller
        const serverMenu = Array.isArray(userData && userData.menu) ? userData.menu : null;
        const menuItems  = (serverMenu && serverMenu.length) ? serverMenu : deriveMenuFromRoles(user.roles || ['Gjest']);
        buildMenu(menuItems);

        // Standard-side
        const firstAction = (menuItems[0] && menuItems[0].action) || 'loadDashboard';
        runMenuAction(firstAction);
    })
    .withFailureHandler(error => {
        setErrorState('Klarte ikke å laste brukerdata: ' + (error && error.message ? error.message : error));
        console.error(error);
        // Fallback for helt offline/feil – vis et minimum
        renderUserInfo({ name:'Gjest', email:'', roles:['Gjest'] });
        const fallbackMenu = deriveMenuFromRoles(['Gjest']);
        buildMenu(fallbackMenu);
        runMenuAction('loadDashboard');
    })
    .uiBootstrap();
}

// ---------- Hjelpefunksjoner for UI ----------
function renderUserInfo(user) {
    if (!user || !user.email) {
        userInfoDiv.innerHTML = '<span>Ikke innlogget</span>';
        return;
    }
    userInfoDiv.innerHTML = `
    <span>Logget inn som: <strong>${user.name || user.email}</strong></span>
    <a href="#" id="logout-btn">Logg ut</a>
    `;
}

function setLoadingState(message) {
    contentArea.innerHTML = `<h2>${message}</h2>`;
}

function setErrorState(message) {
    contentArea.innerHTML = `<h2 style="color: red;">${message}</h2>`;
}

// Bygger meny fra [{name, action}] – action er funksjonsnavn (string)
function buildMenu(menuItems) {
    mainMenuUl.innerHTML = '';

    if (!menuItems || !menuItems.length) {
        mainMenuUl.innerHTML = '<li><em>Ingen menypunkter tilgjengelig</em></li>';
        return;
    }

    menuItems.forEach(item => {
        const li = document.createElement('li');
        const a  = document.createElement('a');
        a.href = '#';
        a.textContent = item.name;
        a.addEventListener('click', e => {
            e.preventDefault();
            document.querySelectorAll('#main-menu a').forEach(el => el.classList.remove('active'));
            a.classList.add('active');
            runMenuAction(item.action);
        });
        li.appendChild(a);
        mainMenuUl.appendChild(li);
    });
}

// Kjør funksjon fra navnet (f.eks. 'loadDashboard')
function runMenuAction(actionName) {
    const fn = typeof actionName === 'string' ? window[actionName] : null;
    if (typeof fn === 'function') {
        fn();
    } else {
        contentArea.innerHTML = `<p>Funksjonen <strong>${actionName}</strong> finnes ikke.</p>`;
        console.warn('Ukjent meny-action:', actionName);
    }
}

// Bygg meny lokalt ut fra roller hvis serveren ikke sendte meny
function deriveMenuFromRoles(roles) {
    roles = Array.isArray(roles) ? roles : ['Gjest'];
    return MENU_CONFIG.filter(item => item.roles.some(r => roles.includes(r)))
    .map(item => ({ name: item.name, action: item.action }));
}

// ---------- “Sider” (kan fylles med ekte data etter hvert) ----------
function loadDashboard() {
    contentArea.innerHTML = '<h2>Dashboard</h2><p>Velkommen til Sameieportalen! Innholdet for dashbordet vil vises her.</p>';
    // Eksempel videre: google.script.run.withSuccessHandler(drawKpi).dashMetrics();
}

function loadAnnouncements() {
    contentArea.innerHTML = '<h2>Oppslag</h2><p>Listen over alle kunngjøringer vil vises her.</p>';
}

function loadMeetings() {
    setLoadingState('Laster møteoversikt…');
    google.script.run
    .withSuccessHandler(meetings => {
        let html = '<h2>Møter & Protokoller</h2>';
        if (!meetings || !meetings.length) {
            html += '<p>Ingen kommende møter funnet.</p>';
        } else {
            html += '<ul>';
            meetings.forEach(meeting => {
                const meetingDate = new Date(meeting.dato).toLocaleDateString('no-NO');
                html += `<li><strong>${meeting.tittel}</strong> - ${meetingDate}</li>`;
            });
            html += '</ul>';
        }
        contentArea.innerHTML = html;
    })
    .withFailureHandler(err => setErrorState('Kunne ikke laste møter: ' + (err && err.message ? err.message : err)))
    .listMeetings_({ scope: 'planned' });
}

function openNewAnnouncementUI() {
    contentArea.insertAdjacentHTML('beforeend','<p>Åpner dialog for nytt oppslag…</p>');
    google.script.run
    .withSuccessHandler(() => setTimeout(() => loadDashboard(), 800))
    .withFailureHandler(err => setErrorState('Kunne ikke åpne dialogen: ' + (err && err.message ? err.message : err)))
    .openNyttOppslagUI();
}

function loadUserAdmin() {
    contentArea.innerHTML = '<h2>Brukeradministrasjon</h2><p>Verktøy for å administrere brukere og roller vil vises her.</p>';
}

function loadJanitorTasks() {
    contentArea.innerHTML = '<h2>Mine Oppgaver</h2><p>Listen over dine tildelte oppgaver som vaktmester vil vises her.</p>';
}

// ---------- Start applikasjonen ----------
document.addEventListener('DOMContentLoaded', init);
</script>
