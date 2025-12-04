<!-- GitHub Copilot Instructions for GMPC-Requisition-Slip -->

# Quick Context

- This repository is a Google Apps Script (GAS) web app that collects requisition requests, creates a Google Doc from a template, generates a PDF, and records rows in a Spreadsheet.
- Server code (GAS) is in `Code.js`. The client UI is the single-page HTML in `index.html`. Project settings live in `appsscript.json`.

# Architecture & Data Flow (short)

- Client (`index.html`) collects the form and signature (canvas → base64) and calls `google.script.run.saveAndCreatePdf(data)`.
- Server (`Code.js`) handles `saveAndCreatePdf(data)`:
  - Writes rows into the spreadsheet identified by `SHEET_FILE_ID`.
  - Makes a copy of the Google Doc template (`TEMPLATE_ID`), replaces placeholders, inserts items and signature, converts to PDF, stores it in `FOLDER_ID`, and returns the public PDF URL.
  - Schedules `generateMasterSummary` via a time-based trigger (best-effort).

# Important Files & Conventions

- `Code.js`: single canonical server file — keep constants at the top: `TEMPLATE_ID`, `FOLDER_ID`, `SHEET_FILE_ID`, `REDIRECT_URL`.
- `index.html`: full client UI and logic. The item catalog appears twice: `ITEM_LIST` in `Code.js` and `itemCategories` in `index.html`. Keep them in sync when adding/removing items.
- `appsscript.json`: defines advanced services (Drive, Sheets, Gmail), runtime (`V8`) and webapp settings (`executeAs: USER_DEPLOYING`, `access: ANYONE_ANONYMOUS`).

# Developer Workflows

- Development uses `clasp` (Google Apps Script CLI). Common commands:
  - `clasp login` — authenticate CLI.
  - `clasp push` — push local files to the Apps Script project (used by this repo).
  - `clasp pull` — fetch remote script files.
  - `clasp open` — open the project in the online editor/UI.
  - `clasp deploy --version <n>` or `clasp deploy` — publish a new deployment (if needed).
- After `clasp push`, test by opening the webapp URL (or use `clasp open` and run `doGet`/preview). Use browser devtools to inspect client-side errors.

# Project-specific Patterns & Gotchas

- Signature handling: the canvas is converted to a base64 PNG in `#sigData` and sent as `requested_by_signature`. Server expects base64 and inserts the image into the doc (see `createPdfFromTemplate`). If altering signature code, keep base64 format compatibility.
- Items/Units: `ITEM_LIST` (server) and `itemCategories` (client mapping to `unitPlural`) must be edited together. Failing to update both causes mismatched units or missing item IDs in sheet rows.
- Real-time stocks: `getRunningStocks()` reads the sheet named `REAL-TIME STOCKS` in the spreadsheet `SHEET_FILE_ID`. If you change sheet structure, update parsing logic there and the client `buildItemsModalList()` accordingly.
- Template placeholders: the template uses tags like `{{BRANCH}}`, `{{ITEMS}}`, `{{REQUESTED_BY_SIGNATURE}}`. Keep placeholder names consistent when modifying the template or the code that replaces them.
- Performance: server code batches sheet appends (fast bulk `setValues`) and minimizes API calls. Preserve that pattern when changing how rows are written.

# Security & Permissions

- `appsscript.json` enables Drive and Sheets advanced services. Changing `SHEET_FILE_ID`, `TEMPLATE_ID`, or `FOLDER_ID` affects the runtime environment and permissions — ensure the deploying user has access.
- Webapp is configured `access: ANYONE_ANONYMOUS` and `executeAs: USER_DEPLOYING`. Be careful when changing these — you may need different OAuth scopes.

# Examples / Quick Edits

- To update the Google Doc template used for PDFs: edit `TEMPLATE_ID` in `Code.js` and ensure the new template contains existing placeholders (e.g., `{{ITEMS}}`, `{{REQUESTED_BY_SIGNATURE}}`).
- To change branches/options shown to users: edit the `<select id="branch">` options in `index.html`.
- To add a new item type: add an entry in `ITEM_LIST` (in `Code.js`) and the matching key/value to `itemCategories['OFFICE SUPPLIES']` in `index.html`, then verify unit mapping in `unitPlural`.

# Testing Tips

- Use browser devtools console to see client-side exceptions (the UI uses `google.script.run.withFailureHandler` to alert errors from server calls).
- On server-side debugging, add `Logger.log(...)` in `Code.js` and check logs in the Apps Script editor or via the Stackdriver/Cloud Logs in the GCP project associated.

# When in doubt — where to look

- Primary server logic: `Code.js` (PDF creation, sheet writes, triggers).
- Primary client logic: `index.html` (form, signature pad, item selection, `google.script.run` calls).
- Deployment/config: `appsscript.json` and clasp commands.

# Feedback request

Please review this guidance and tell me if you'd like any section expanded (deploy steps, permission troubleshooting, or template editing examples).
