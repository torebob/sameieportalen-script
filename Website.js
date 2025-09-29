<script>
document.addEventListener('DOMContentLoaded', function() {
  // Load the home page by default
  loadPage('home');

  document.querySelector('nav').addEventListener('click', function(e) {
    if (e.target.tagName === 'A') {
      e.preventDefault();
      const pageId = e.target.dataset.page;
      loadPage(pageId);
    }
  });
});

function loadPage(pageId) {
    const contentDiv = document.getElementById('content');
    contentDiv.innerHTML = '<p>Laster inn...</p>';

    if (pageId === 'news') {
        google.script.run.withSuccessHandler(renderNews).getNewsFeed();
    } else if (pageId === 'docs') {
        google.script.run.withSuccessHandler(renderDocuments).getDocuments();
    } else {
        // First, try to get the page without a password.
        google.script.run
            .withSuccessHandler(function(page) {
                if (page && page.authRequired) {
                    // If auth is required, render the password form.
                    renderPasswordForm(pageId);
                } else if (page) {
                    // If page is public or password was correct, render it.
                    renderPageContent(pageId, page);
                } else {
                    // Handle case where page is not found
                    contentDiv.innerHTML = '<h2>Velkommen!</h2><p>Dette er hjemmesiden for sameiet. Bruk menyen over for å navigere.</p>';
                }
            })
            .withFailureHandler(function(error) {
                contentDiv.innerHTML = `<p>En feil oppstod: ${error.message}</p>`;
            })
            .getPageContent(pageId, null); // Always call without password first
    }
}

function renderPageContent(pageId, page) {
    const contentDiv = document.getElementById('content');
    let editButtonHtml = `<br><br><button onclick="editPage('${pageId}')">Rediger Side</button>`;
    contentDiv.innerHTML = `<h2>${page.title}</h2><div>${page.content}</div>${editButtonHtml}`;
}

function renderPasswordForm(pageId) {
    const contentDiv = document.getElementById('content');
    contentDiv.innerHTML = `
        <h2>Passordbeskyttet Side</h2>
        <p>Denne siden krever et passord for å vises.</p>
        <form id="password-form">
            <input type="password" id="password-input" placeholder="Skriv inn passord">
            <button type="submit">Lås Opp</button>
        </form>
        <p id="password-error" style="color: red;"></p>
    `;

    document.getElementById('password-form').addEventListener('submit', function(e) {
        e.preventDefault();
        const password = document.getElementById('password-input').value;
        google.script.run
            .withSuccessHandler(function(response) {
                if (response && response.ok === false) {
                    document.getElementById('password-error').innerText = 'Feil passord. Prøv igjen.';
                } else {
                    renderPageContent(pageId, response);
                }
            })
            .withFailureHandler(function(error) {
                document.getElementById('password-error').innerText = `En feil oppstod: ${error.message}`;
            })
            .verifyPassword(pageId, password);
    });
}

function editPage(pageId) {
    // Construct the URL for the editor using the global variable
    const editorUrl = `${APP_URL}?action=edit&page=${pageId}`;
    // Open the editor in a new tab or modal. A new tab is simpler.
    window.open(editorUrl, '_blank');
}

function renderNews(articles) {
  const contentDiv = document.getElementById('content');
  if (!articles || articles.length === 0) {
    contentDiv.innerHTML = '<h2>Nyheter</h2><p>Ingen nyheter å vise.</p>';
    return;
  }
  let html = '<h2>Nyheter</h2>';
  articles.forEach(article => {
    html += `
      <article class="news-article">
        <h3>${article.title}</h3>
        <p>${article.content}</p>
        <small>Publisert: ${new Date(article.publishedDate).toLocaleDateString()}</small>
      </article>
    `;
  });
  contentDiv.innerHTML = html;
}

function renderDocuments(documents) {
  const contentDiv = document.getElementById('content');
  if (!documents || documents.length === 0) {
    contentDiv.innerHTML = '<h2>Dokumenter</h2><p>Ingen dokumenter tilgjengelig.</p>';
    return;
  }
  let html = '<h2>Dokumenter</h2><ul>';
  documents.forEach(doc => {
    html += `
      <li>
        <a href="${doc.url}" target="_blank">${doc.title}</a>
        <p>${doc.description || ''}</p>
      </li>
    `;
  });
  html += '</ul>';
  contentDiv.innerHTML = html;
}
</script>