/**
 * Dashboard Admin (v1.0.0)
 * Handles the creation and logic for the admin-only sidebar.
 */

/**
 * Determines which dashboard to show based on user role.
 */
function openDashboardAuto() {
  return getCurrentUserInfo().isAdmin ? openDashboard() : openDashboardModal();
}

/**
 * Opens the user-facing modal dashboard.
 */
function openDashboardModal() {
  dashOpen('DASHBOARD_HTML');
}

/**
 * Builds and displays the admin sidebar.
 */
function openDashboard() {
  const functionName = 'openDashboard';
  try {
    const user = getCurrentUserInfo();
    const html = _buildAdminSidebarHtml_(user);
    const output = HtmlService.createHtmlOutput(html)
      .setTitle("Admin Dashboard")
      .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(output);
  } catch (e) {
    Logger.error(functionName, 'Failed to build and show admin sidebar.', { errorMessage: e.message });
    SpreadsheetApp.getUi().alert('Could not open admin dashboard.');
  }
}

/**
 * Generates the complete HTML for the admin sidebar.
 * @private
 */
function _buildAdminSidebarHtml_(user) {
  const appName = "Sameieportalen"; // Or fetch from a config
  const appVersion = "2.4.0";
  
  // You can move the CSS and JS into separate .html files and include them
  // for even better organization, but this works too.
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          /* All your CSS from the old file goes here */
          body { font-family: sans-serif; padding: 10px; }
          button { margin-top: 5px; }
        </style>
      </head>
      <body>
        <h3>${appName} Admin</h3>
        <p>Logget inn som: ${user.email}</p>
        <hr>
        <h4>Admin Actions</h4>
        <button onclick="runAdminAction('runAllChecks')">Kjør Systemsjekk</button>
        <button onclick="runAdminAction('adminEnableDevTools')">Aktiver Utviklerverktøy</button>
        <div id="status"></div>
        
        <script>
          function runAdminAction(actionName) {
            document.getElementById('status').textContent = 'Kjører...';
            google.script.run
              .withSuccessHandler(response => {
                document.getElementById('status').textContent = 'Suksess: ' + response;
              })
              .withFailureHandler(error => {
                document.getElementById('status').textContent = 'Feil: ' + error.message;
              })
              [actionName](); // Dynamically call the function
          }
        </script>
      </body>
    </html>
  `;
}