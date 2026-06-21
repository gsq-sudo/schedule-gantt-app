# Schedule Gantt

A static schedule management app that uses the same Google Drive API pattern as the earlier Project Tracker: local-first browser storage with sync to Google Drive's hidden `appDataFolder`.

## Run Locally

```powershell
python -m http.server 4173
```

Open `http://localhost:4173`.

## Google Drive Setup

Use the same Google Cloud OAuth setup as the earlier app:

1. Enable the Google Drive API.
2. Create or reuse an OAuth Client ID for a web application.
3. Add the app origin, such as `http://localhost:4173`.
4. Put the client ID in `config.js`.

The app requests only:

```text
https://www.googleapis.com/auth/drive.appdata
```

Schedule data is stored in one JSON file named `schedule-gantt-data.json` in the Drive app data folder.

## Schedule Behavior

- Project bars can be recolored from an Office-style color palette.
- New projects default into the timeline group for their assigned owner when that owner already has timeline rows.
- Projects can be reordered on the Gantt timeline by dragging the grip handle in the row label; the page auto-scrolls near the top and bottom edges during reorder.
- Completing a project archives it; archived projects remain visible on the timeline with distinct styling and can be restored.
- The Gantt toolbar can hide completed projects from the timeline while keeping them in the Projects list.
- Weekends are visible on the timeline but do not count as workdays or utilization.
- Utilization above 100% is shown in red.
- The PDF export button opens a print-ready report with segmented Gantt pages and the full project list; choose "Save as PDF" in the print dialog to download it.

## Files

- `index.html`
- `styles.css`
- `app.js`
- `config.js`

There is no build step. Upload the files to any static host whose origin is allowed by the OAuth client.
