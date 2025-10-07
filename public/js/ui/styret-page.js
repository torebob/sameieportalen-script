document.addEventListener('DOMContentLoaded', () => {
    // Henter elementene fra HTML-filen
    const sendKnapp = document.getElementById('send-kunngjoring-knapp');
    const tittelFelt = document.getElementById('kunngjoring-tittel');
    const innholdFelt = document.getElementById('kunngjoring-innhold');
    const maalgruppeFelt = document.getElementById('kunngjoring-maalgruppe');
    const statusMelding = document.getElementById('status-melding');

    // Sjekker at alle elementene finnes
    if (!sendKnapp || !tittelFelt || !innholdFelt || !maalgruppeFelt || !statusMelding) {
        console.error('Kunne ikke finne et eller flere av skjema-elementene på siden.');
        return;
    }

    // Legger til en lytter på knappen som kjører når den klikkes
    sendKnapp.addEventListener('click', () => {
        const payload = {
            tittel: tittelFelt.value.trim(),
                               innhold: innholdFelt.value.trim(),
                               maalgruppe: maalgruppeFelt.value
        };

        if (!payload.tittel || !payload.innhold) {
            statusMelding.textContent = 'Feil: Du må fylle ut både tittel og melding.';
            statusMelding.style.color = 'red';
            return;
        }

        // Deaktiverer knappen og viser en melding mens vi sender
        sendKnapp.disabled = true;
        sendKnapp.textContent = 'Sender...';
        statusMelding.textContent = 'Sender kunngjøring...';
        statusMelding.style.color = 'black';

        // Dette er kallet til din Google Apps Script backend!
        google.script.run
        .withSuccessHandler(response => {
            if (response.ok) {
                statusMelding.textContent = response.message;
                statusMelding.style.color = 'green';
                tittelFelt.value = '';
                innholdFelt.value = '';
            } else {
                statusMelding.textContent = 'En feil oppstod: ' + (response.error || response.message);
                statusMelding.style.color = 'red';
            }
            sendKnapp.disabled = false;
            sendKnapp.textContent = 'Send kunngjøring';
        })
        .withFailureHandler(error => {
            statusMelding.textContent = 'Teknisk feil: ' + error.message;
            statusMelding.style.color = 'red';
            console.error('Feil ved kall til sendOppslag:', error);
            sendKnapp.disabled = false;
            sendKnapp.textContent = 'Send kunngjøring';
        })
        .sendOppslag(payload); // Her kalles din 'sendOppslag'-funksjon i Apps Script
    });
});
