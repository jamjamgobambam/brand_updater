// =============================================================================
// main.js — Batch entry points and trigger functions
// Depends on globals from utils.js and slides-updater.js
// =============================================================================

/**
 * Runs the full brand update pipeline on every Google Slides presentation
 * found directly in the specified Drive folder.
 *
 * Apps Script execution limits:
 *   - 6 minutes for free accounts; 30 minutes for Google Workspace accounts.
 * If the folder is large and the script times out, run it on smaller subfolders.
 *
 * @param {string}  folderId  Google Drive folder ID.
 * @param {boolean} [dryRun=false]  Passed through to replaceLogos for each file.
 */
function updateAllSlidesInFolder(folderId, dryRun) {
  const SLIDES_MIME = "application/vnd.google-apps.presentation";
  const folder = DriveApp.getFolderById(folderId);
  const files  = folder.getFiles();

  let processed = 0;
  let failed    = 0;

  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== SLIDES_MIME) continue;

    const fileName = file.getName();
    Logger.log("Processing: %s (%s)", fileName, file.getId());

    try {
      updateSlidesPresentation(file.getId(), dryRun);
      processed++;
      Logger.log("  ✓ Done: %s", fileName);
    } catch (err) {
      failed++;
      Logger.log("  ✗ FAILED: %s — %s", fileName, err.message);
    }
  }

  Logger.log(
    "Batch complete. Processed: %d  Failed: %d",
    processed,
    failed
  );
}

/**
 * Runs the full brand update pipeline on every Google Docs document
 * found directly in the specified Drive folder.
 *
 * @param {string} folderId  Google Drive folder ID.
 */
function updateAllDocsInFolder(folderId) {
  const DOCS_MIME = "application/vnd.google-apps.document";
  const folder = DriveApp.getFolderById(folderId);
  const files  = folder.getFiles();

  let processed = 0;
  let failed    = 0;

  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== DOCS_MIME) continue;

    const fileName = file.getName();
    Logger.log("Processing: %s (%s)", fileName, file.getId());

    try {
      updateDocsDocument(file.getId());
      processed++;
      Logger.log("  ✓ Done: %s", fileName);
    } catch (err) {
      failed++;
      Logger.log("  ✗ FAILED: %s — %s", fileName, err.message);
    }
  }

  Logger.log(
    "Batch complete. Processed: %d  Failed: %d",
    processed,
    failed
  );
}
