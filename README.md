# Valero's log

A small static page (phone-first) for logging hard moments. Entries are saved to one JSON file
(`ValerosLog/valeros-log.json`) in OneDrive via Microsoft Graph. Nothing is stored in a server.

## One-time setup
1. Entra admin center → App registrations → New registration. Supported accounts: *personal Microsoft accounts*
   (or "any directory and personal"). Platform: **Single-page application**, redirect URI = the page URL
   (e.g. `https://<user>.github.io/<repo>/` and `http://localhost:8080/` for testing).
2. API permissions (delegated, Microsoft Graph): `User.Read`, `Files.ReadWrite.All`.
3. Paste the Application (client) ID into `CLIENT_ID` at the top of `app.js`.

## Using it
- First person signs in and taps **Start a new log**. Then use **Get a link to share this log** and send it to the other parent.
- The other parent signs in with their own account and pastes the link once.
- Both phones then read/write the same file. Saves use the file's ETag so two saves never overwrite each other.

The guidance comes from "Managing Severe Preschooler Tantrums.docx"; it isn't medical advice.
