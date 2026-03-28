// =============================================================================
// webapp.js — Google Apps Script web app server
// Handles HTTP GET (doGet), URL parsing, and routing to the brand updaters.
// Depends on globals from utils.js, slides-updater.js, and docs-updater.js.
// =============================================================================

// ---------------------------------------------------------------------------
// Step 11 — doGet
// ---------------------------------------------------------------------------

/**
 * Entry point for HTTP GET requests to the web app URL.
 * Serves the index.html file as a sandboxed HTML page.
 *
 * @returns {HtmlOutput}
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Brand Updater");
}

// ---------------------------------------------------------------------------
// Step 8 — extractIdAndType
// ---------------------------------------------------------------------------

/**
 * Parses any Google Drive / Docs / Slides URL and returns { id, type }.
 *
 * Supported patterns:
 *   docs.google.com/presentation/d/{ID}  → 'slides'
 *   docs.google.com/document/d/{ID}      → 'docs'
 *   drive.google.com/drive/[u/N/]folders/{ID}  → 'folder'
 *   drive.google.com/file/d/{ID}         → 'driveFile'
 *   drive.google.com/open?id={ID}        → 'driveFile'
 *   Unrecognized                         → { id: null, type: 'invalid' }
 *
 * @param {string} url
 * @returns {{ id: string|null, type: string }}
 */
function extractIdAndType(url) {
  if (!url || typeof url !== "string") {
    return { id: null, type: "invalid" };
  }

  var trimmed = url.trim();

  // Slides presentation
  var slidesMatch = trimmed.match(/docs\.google\.com\/presentation\/d\/([a-zA-Z0-9_-]+)/);
  if (slidesMatch) return { id: slidesMatch[1], type: "slides" };

  // Docs document
  var docsMatch = trimmed.match(/docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/);
  if (docsMatch) return { id: docsMatch[1], type: "docs" };

  // Drive folder (supports /drive/folders/, /drive/u/0/folders/, etc.)
  var folderMatch = trimmed.match(/drive\.google\.com\/drive(?:\/u\/\d+)?\/folders\/([a-zA-Z0-9_-]+)/);
  if (folderMatch) return { id: folderMatch[1], type: "folder" };

  // Drive file by path
  var driveFileMatch = trimmed.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (driveFileMatch) return { id: driveFileMatch[1], type: "driveFile" };

  // Drive open?id= link
  var openIdMatch = trimmed.match(/drive\.google\.com\/open\?(?:[^&]*&)*id=([a-zA-Z0-9_-]+)/);
  if (!openIdMatch) {
    // Also handle id= as first param
    openIdMatch = trimmed.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/);
  }
  if (openIdMatch) return { id: openIdMatch[1], type: "driveFile" };

  return { id: null, type: "invalid" };
}

// ---------------------------------------------------------------------------
// Step 9 — resolveFileType
// ---------------------------------------------------------------------------

/**
 * Resolves the type of a generic Drive file link by checking its MIME type.
 * Used only when extractIdAndType returns 'driveFile'.
 *
 * @param {string} id  Google Drive file ID.
 * @returns {'slides'|'docs'|'unsupported'}
 */
function resolveFileType(id) {
  var mimeType = DriveApp.getFileById(id).getMimeType();
  if (mimeType === "application/vnd.google-apps.presentation") return "slides";
  if (mimeType === "application/vnd.google-apps.document")     return "docs";
  return "unsupported";
}

// ---------------------------------------------------------------------------
// Step 10 — processUrl
// ---------------------------------------------------------------------------

/**
 * Server-side function called from the client via google.script.run.
 * Routes to the appropriate updater based on the parsed URL type.
 *
 * For 'folder' type, iterates all files in the folder and routes each by
 * MIME type. Only Slides and Docs files are processed; others are skipped.
 *
 * @param {string}  url     Any Google Drive / Docs / Slides URL.
 * @param {boolean} dryRun  When true, logo replacement is previewed only.
 * @returns {{ processed: number, failed: number, details: Object[] }}
 */
function processUrl(url, dryRun) {
  var parsed = extractIdAndType(url);

  if (parsed.type === "invalid") {
    return {
      processed: 0,
      failed: 0,
      details: [{ name: url, status: "failed", error: "Unrecognized URL format. Please paste a Google Slides, Docs, or Drive folder/file URL." }],
    };
  }

  var isDryRun = dryRun === true;

  // For a single file type, wrap it in a one-item list to share the loop below
  var items = [];

  if (parsed.type === "folder") {
    var folder = DriveApp.getFolderById(parsed.id);
    var files   = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var mime = file.getMimeType();
      var itemType;
      if (mime === "application/vnd.google-apps.presentation") {
        itemType = "slides";
      } else if (mime === "application/vnd.google-apps.document") {
        itemType = "docs";
      } else {
        continue; // skip non-Slides/Docs files
      }
      items.push({ id: file.getId(), name: file.getName(), type: itemType });
    }
  } else {
    var resolvedType = parsed.type;
    if (resolvedType === "driveFile") {
      resolvedType = resolveFileType(parsed.id);
    }
    var fileName;
    try {
      fileName = DriveApp.getFileById(parsed.id).getName();
    } catch (e) {
      fileName = parsed.id;
    }
    if (resolvedType === "unsupported") {
      return {
        processed: 0,
        failed: 0,
        details: [{ name: fileName, status: "failed", error: "File type is not supported. Only Google Slides and Docs files can be updated." }],
      };
    }
    items.push({ id: parsed.id, name: fileName, type: resolvedType });
  }

  var processed = 0;
  var failed    = 0;
  var details   = [];

  items.forEach(function(item) {
    try {
      if (item.type === "slides") {
        updateSlidesPresentation(item.id, isDryRun);
      } else if (item.type === "docs") {
        updateDocsDocument(item.id);
      }
      processed++;
      details.push({ name: item.name, status: "ok" });
    } catch (err) {
      failed++;
      details.push({ name: item.name, status: "failed", error: err.message });
    }
  });

  return { processed: processed, failed: failed, details: details };
}
