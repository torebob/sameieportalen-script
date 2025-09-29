# Setup Instructions for Gjøremål (Task Management) Module

Follow these steps to configure the backend for the task management application.

## Step 1: Create the Google Sheet Database

1.  **Create a new Google Sheet.** You can do this by visiting [sheet.new](https://sheet.new).
2.  **Name the spreadsheet.** Give it a clear name, for example, "Styret Portalen Database".
3.  **Get the Spreadsheet ID.** The ID is the long string of characters in the URL between `/d/` and `/edit`.
    -   Example URL: `https://docs.google.com/spreadsheets/d/1aBcDeFgHiJkLmNoPqRsTuVwXyZa_BcDeFgHiJkLmNoP/edit`
    -   In this example, the ID is `1aBcDeFgHiJkLmNoPqRsTuVwXyZa_BcDeFgHiJkLmNoP`.
4.  **Paste the ID into the script.** Open the `Gjoremal_Backend.gs` file and replace `YOUR_SHEET_ID_HERE` with the ID you just copied.

    ```javascript
    // BEFORE
    const DB_SHEET_ID = 'YOUR_SHEET_ID_HERE';

    // AFTER
    const DB_SHEET_ID = '1aBcDeFgHiJkLmNoPqRsTuVwXyZa_BcDeFgHiJkLmNoP';
    ```

5.  **Create the `Tasks` sheet.**
    -   Rename the default "Sheet1" to `Tasks`.
    -   Add the following headers in the first row, exactly as written:
        `id`, `title`, `description`, `assignee`, `dueDate`, `status`, `attachmentUrl`

6.  **Create the `Users` sheet.**
    -   Click the `+` icon at the bottom left to add a new sheet.
    -   Rename this new sheet to `Users`.
    -   Add the following headers in the first row: `name`, `email`.
    -   Populate this sheet with the names and email addresses of the board members who can be assigned tasks.

7.  **Create the `Suppliers` sheet.**
    -   Click the `+` icon to add another new sheet.
    -   Rename it to `Suppliers`.
    -   Add the following headers in the first row, exactly as written:
        `id`, `name`, `contactPerson`, `phone`, `email`, `contractId`, `services`, `notes`

## Step 2: Create the Google Drive Folder for Attachments

1.  **Create a new folder in Google Drive.** You can do this by visiting [drive.google.com](https://drive.google.com) and clicking "New" > "Folder".
2.  **Name the folder.** Give it a clear name, for example, "Portalen Vedlegg".
3.  **Get the Folder ID.** The ID is the string of characters at the end of the folder's URL.
    -   Example URL: `https://drive.google.com/drive/folders/2bCdEfGhIjKlMnOpQrStUvWxYzAbCdEfG`
    -   In this example, the ID is `2bCdEfGhIjKlMnOpQrStUvWxYzAbCdEfG`.
4.  **Paste the ID into the script.** Open `Gjoremal_Backend.gs` and replace `YOUR_FOLDER_ID_HERE` with the folder ID.

    ```javascript
    // BEFORE
    const ATTACHMENTS_FOLDER_ID = 'YOUR_FOLDER_ID_HERE';

    // AFTER
    const ATTACHMENTS_FOLDER_ID = '2bCdEfGhIjKlMnOpQrStUvWxYzAbCdEfG';
    ```

## Step 3: Grant Permissions

When you first run the application, Google will ask you to grant permissions for the script to access your Google Sheets and Google Drive. You must approve these permissions for the application to function correctly.

Once you have completed these steps, the application should be fully configured and operational.