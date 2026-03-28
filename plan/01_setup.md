# Plan Part 1: Google Apps Script Project Setup (clasp + VS Code)

The standard approach is **clasp** — Google's official CLI that syncs local files with Apps Script projects.

## Decisions

- Language: JavaScript (no TypeScript build step)
- Project type: New standalone Apps Script project
- Binding: Standalone (not attached to a Drive file)

---

## Phase 1 — Prerequisites

1. **Verify Node.js is installed** — run `node -v` and `npm -v` in a terminal. If missing, install from https://nodejs.org.
2. **Enable the Apps Script API** — go to https://script.google.com/home/usersettings and toggle the API on (required for clasp).

---

## Phase 2 — Initialize the Project

1. **`npm init -y`** — creates `package.json` in your workspace.
2. **`npm install -g @google/clasp`** — installs clasp globally.
3. **`npm install --save-dev @types/google-apps-script`** — installs type definitions for Apps Script IntelliSense in VS Code.
4. **`clasp login`** — opens a browser for Google OAuth.
5. **`clasp create --title "brand_updater" --type standalone`** — creates the project in Google Drive and generates:
   - `.clasp.json` — contains the `scriptId` linking this folder to the cloud project
   - `appsscript.json` — the Apps Script manifest

---

## Phase 3 — Configure VS Code Support

1. **Create `jsconfig.json`** — tells VS Code to use `@types/google-apps-script` for IntelliSense on globals like `SlidesApp`, `DocumentApp`, `DriveApp`, etc.
2. **Create `.claspignore`** — prevents clasp from pushing `node_modules/`, `package.json`, `jsconfig.json`, etc. to the cloud.
3. **Create `.gitignore`** — excludes `node_modules/` and `.clasp.json` from source control.

---

## Phase 4 — Day-to-Day Workflow

| Command | Purpose |
|---|---|
| `clasp push` | Upload local changes to Apps Script cloud |
| `clasp pull` | Download remote changes |
| `clasp open-script` | Open the project in the browser editor |

---

## Files to Create

- `package.json` (via npm init)
- `jsconfig.json` — references @types/google-apps-script for IntelliSense
- `.claspignore` — excludes node_modules/, package*.json, jsconfig.json
- `.gitignore` — excludes node_modules/, .clasp.json
- `appsscript.json` — Apps Script manifest (auto-created by clasp create)
- `.clasp.json` — clasp config with scriptId (auto-created)
- `Code.js` — main script file (initial placeholder)

---

## Verification

1. Run `clasp open-script` to confirm the project appears in the Apps Script editor.
2. Add a test function to `Code.js`, run `clasp push`, and verify it appears in the browser editor.
3. Run the function from the browser editor to confirm execution works end-to-end.
