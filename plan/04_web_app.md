# Plan Part 4: Web App — Docs Updater + Web App UI

## Summary

Deploy the brand updater as a Google Apps Script web app so users can paste any
Google Drive folder or file URL (Slides or Docs) and trigger the brand update
pipeline from a browser. Includes implementing the missing `docs-updater.js` as
a prerequisite.

**Decisions:**

- Docs updater is implemented in full (not deferred)
- Access: Anyone in Google Workspace org
- Dry-run checkbox (checked by default for safety)

---

## Phase 1 — Docs Updater (`docs-updater.js`)

*Prerequisite for docs URL support in the web app.*

### Step 1 — Add Docs advanced service to `appsscript.json`

Add to `enabledAdvancedServices`:

```json
{ "userSymbol": "Docs", "serviceId": "docs", "version": "v1" }
```

Also enable it in the Apps Script cloud editor → **Services panel** (manual step — required alongside the manifest entry).

### Step 2 — `traverseContentArray(contentArray, callback)`

Recursively walks a Docs content array (body, header, or footer). Content items
can be paragraphs or tables; tables contain rows → cells, each with their own
content array. Calls `callback({ startIndex, endIndex, style, namedStyleType })` for every
`textRun`. Note the function is named `traverseContentArray` in the implementation (not `traverseContent`).

### Step 3 — `buildDocColorRequests(contentArrays, colorMap)`

Iterates all content sources (body + all headers + all footers). For each
`textRun`, checks `foregroundColor.color.rgbColor` against `colorMap` using the
existing `normalizedRgbMatches` helper from `utils.js`. Returns an array of
`updateTextStyle` request objects.

### Step 4 — `buildDocFontRequests(contentArrays, fontMap)`

Same traversal as Step 3. Checks `weightedFontFamily.fontFamily` (preferred) or
`fontFamily` on each `textRun.textStyle`. Skips `null` font family (inheriting
from Named Style — leave untouched). Returns `updateTextStyle` request objects
preserving font weight.

### Step 5 — `buildDocLogoRequests(doc, newLogoUrl, dryRun)`

Implemented as described in Plan 03 Step 11 (delete + re-insert pairs). Finds
all logo inline objects matching `LOGO_CONFIG.docsLogo` criteria, builds a
`deleteContentRange` + `insertInlineImage` pair per match, and returns them
sorted in reverse index order.

The `newLogoUrl` parameter must be a direct, publicly accessible image URL
(not a Drive redirect URL). Set `LOGO_CONFIG.docsLogo.newLogoUrl` before running.

### Step 6 — `updateDocsDocument(docId, dryRun)`

Orchestrator function for a single document (see Plan 03 Step 13 for full detail). The `dryRun` parameter is passed through to `replaceDocLogos`.

### Step 7 — `updateAllDocsInFolder(folderId)` (in `main.js`)

Mirrors `updateAllSlidesInFolder`. Iterates the folder, filters for MIME type
`application/vnd.google-apps.document`, calls `updateDocsDocument(file.getId())`
for each, and logs processed/failed counts.

---

## Phase 2 — Web App Server (`webapp.js`)

### Step 8 — `extractIdAndType(url)`

Parses any Google URL and returns `{ id, type }`. Supported patterns:

| Input URL pattern | Returned type |
|---|---|
| `docs.google.com/presentation/d/{ID}` | `'slides'` |
| `docs.google.com/document/d/{ID}` | `'docs'` |
| `drive.google.com/drive/[u/N/]folders/{ID}` | `'folder'` |
| `drive.google.com/file/d/{ID}` | `'driveFile'` |
| `drive.google.com/open?id={ID}` | `'driveFile'` |
| Unrecognized | `{ id: null, type: 'invalid' }` |

### Step 9 — `resolveFileType(id)`

Used only for `'driveFile'` type (generic Drive links that don't reveal the
MIME type in the URL). Calls `DriveApp.getFileById(id).getMimeType()` and maps:

- `application/vnd.google-apps.presentation` → `'slides'`
- `application/vnd.google-apps.document` → `'docs'`
- Anything else → `'unsupported'`

### Step 10 — `processUrl(url, dryRun)`

The server-side function called from the client via `google.script.run`. Routes
to the appropriate updater based on `extractIdAndType`:

- `'folder'` — iterate files with `DriveApp`, route each by MIME type
- `'slides'` — `updateSlidesPresentation(id, dryRun)`
- `'docs'` — `updateDocsDocument(id)`
- `'driveFile'` — `resolveFileType(id)` then route as above
- `'invalid'` — return an error immediately

Returns:
```js
{
  processed: number,
  failed: number,
  details: [{ name: string, status: 'ok' | 'failed', error?: string }]
}
```

### Step 11 — `doGet()`

Entry point for HTTP GET requests to the web app URL. Serves `index.html`:

```js
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Brand Updater');
}
```

> **Note:** The planned `.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.SAMEORIGIN)` clickjacking protection is not in the current implementation. It can be added as a follow-up.

---

## Phase 3 — Web App UI (`index.html`)

### Step 12 — Create `index.html`

Single-page form with:

- **URL input** — accepts any Google Drive / Docs / Slides URL; placeholder
  text shows example URL formats
- **Dry-run checkbox** — checked by default; label explains it will preview
  changes without modifying files
- **Update button** — disabled while a run is in progress
- **Loading indicator** — visible during the `google.script.run` call
- **Results panel** — shows processed/failed counts and a scrollable per-file
  detail list (name, status, error message if failed)
- **Error handler** — catches script-level failures (e.g. permission error,
  timeout) and shows a human-readable message

Wiring pattern:
```js
google.script.run
  .withSuccessHandler(showResults)
  .withFailureHandler(showError)
  .processUrl(url, dryRun);
```

---

## Phase 4 — Manifest & Deployment

### Step 13 — Update `appsscript.json`

Add the Docs advanced service entry (same structure as the existing Slides
entry). After `clasp push`, also enable it in the Services panel — both steps
are required.

### Step 14 — `clasp push`

Uploads all new and changed files:

- `docs-updater.js` (new)
- `webapp.js` (new)
- `index.html` (new)
- `appsscript.json` (updated)
- `main.js` (updated)

### Step 15 — Deploy as Web App

**First deployment** — access settings must be configured in the browser editor once:

1. Open the project: `clasp open-script`
2. **Deploy → New deployment**
3. Type: **Web app**
4. Execute as: **Me**
5. Who has access: **Anyone within [org] (Google Workspace)**
6. Click **Deploy** and copy the web app URL

**Alternatively (and for all subsequent updates), deploy entirely from the terminal:**

```bash
# Create an immutable version snapshot
clasp version "description"
# Note the version number printed (e.g., 1)

# Deploy that version
clasp deploy 1 "web app v1"
# Prints the deploymentId and web app URL
```

For subsequent updates after the first deployment:

```bash
clasp push
clasp version "description"
clasp redeploy <deploymentId> --versionNumber <newVersion> --description "description"
```

To list existing deployments and their IDs:

```bash
clasp deployments
```

To open the live web app URL in a browser:

```bash
clasp open-web-app
```

> **Prerequisite for `clasp open-web-app`:** The `appsscript.json` manifest must
> include a `webapp` block, otherwise clasp reports "No web app entry point found":
>
> ```json
> "webapp": {
>   "executeAs": "USER_DEPLOYING",
>   "access": "DOMAIN"
> }
> ```
>
> Valid values: `executeAs` — `USER_DEPLOYING` (Me) or `USER_ACCESSING` (user running the app).
> `access` — `MYSELF`, `DOMAIN`, or `ANYONE`.

> **Note:** The `execute as` / `who has access` settings can only be configured
> in the browser editor. Run `clasp open-script` and set them once during the
> first deployment; all future redeployments via `clasp redeploy` will inherit
> those settings.

---

## Files Changed

| File | Action |
|---|---|
| `docs-updater.js` | Create |
| `webapp.js` | Create |
| `index.html` | Create |
| `appsscript.json` | Update — add Docs service + `script.external_request` OAuth scope |
| `main.js` | Update — add `updateAllDocsInFolder()` |
| `utils.js` | Update — add `docsLogo` key to `LOGO_CONFIG` |
| `slides-updater.js` | No changes — reused as-is |

---

## Verification

1. `clasp push` — confirm no errors in the push output
2. Load web app URL → form renders correctly
3. Paste a **Slides URL**, dry-run **ON** → results list appears, presentation
   is unchanged
4. Paste a **Slides URL**, dry-run **OFF** → colors, fonts updated in the
   actual presentation
5. Paste a **Docs URL**, dry-run **ON** → traversal returns expected counts
6. Paste a **Drive folder URL** → all Slides and Docs inside are processed
7. Paste an **invalid URL** → graceful error message rendered in the UI
8. Run `logAllImages()` on a representative presentation to verify logo
   position thresholds before any live logo replacement

---

## Known Limitations

1. **Execution timeout** — Apps Script web app requests time out at 6 min (free
   accounts) / 30 min (Google Workspace). For large folders the run may be
   interrupted. A follow-up improvement would split processing across
   time-based triggers.

2. **Dry-run has no effect on Docs logo replacement from the web app** —
   `processUrl` in `webapp.js` calls `updateDocsDocument(item.id)` without
   passing `isDryRun`, so the dry-run checkbox only affects Slides logo
   replacement. Docs colors and fonts are always applied regardless of the
   checkbox (they have no dry-run mode). To enable dry-run for Docs logos
   via the web app, change the call to
   `updateDocsDocument(item.id, isDryRun)`.

3. **UI dynamic mode indicator** — When the dry-run checkbox is unchecked,
   the button label changes from "Preview Changes" to "Apply Changes" and
   an amber warning reads "Live mode — matching logos will be replaced in
   your files. This cannot be undone." This behavior is not described in
   the original plan but is fully implemented in `index.html`.
