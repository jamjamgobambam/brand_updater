// =============================================================================
// docs-updater.js — Google Docs brand updater (colors, fonts, logos)
// Depends on globals defined in utils.js: COLOR_MAP, FONT_MAP, LOGO_CONFIG,
// normalizedRgbMatches, hexToNormalizedRgb, driveFileUrl
// Requires the Docs Advanced Service (userSymbol: "Docs") enabled in the
// Apps Script project and appsscript.json.
// =============================================================================

// ---------------------------------------------------------------------------
// Step 2 — traverseContent
// ---------------------------------------------------------------------------

/**
 * Recursively walks a Docs content array (body, header, or footer).
 * Visits every textRun in paragraphs and tables (including nested cells).
 * Calls callback({ startIndex, endIndex, style }) for each textRun found.
 *
 * @param {Object[]} contentArray  The .content array of a Body, Header, or Footer.
 * @param {Function} callback      Called with each textRun's { startIndex, endIndex, style }.
 */
function traverseContent(contentArray, callback) {
  if (!contentArray) return;

  contentArray.forEach(function(structuralElement) {
    if (structuralElement.paragraph) {
      (structuralElement.paragraph.elements || []).forEach(function(element) {
        if (!element.textRun) return;
        callback({
          startIndex: element.startIndex,
          endIndex:   element.endIndex,
          style:      element.textRun.textStyle || {},
        });
      });
    }

    if (structuralElement.table) {
      (structuralElement.table.tableRows || []).forEach(function(row) {
        (row.tableCells || []).forEach(function(cell) {
          traverseContent(cell.content, callback);
        });
      });
    }
  });
}

// ---------------------------------------------------------------------------
// Step 3 — buildDocColorRequests
// ---------------------------------------------------------------------------

/**
 * Iterates all content sources (body + headers + footers) and builds
 * updateTextStyle request objects for every textRun whose foreground color
 * matches an entry in colorMap.
 *
 * @param {Object[]} contentArrays  Array of .content arrays to traverse.
 * @param {Object[]} colorMap       Array of { oldHex, newHex } entries (COLOR_MAP).
 * @returns {Object[]}              Array of updateTextStyle request objects.
 */
function buildDocColorRequests(contentArrays, colorMap) {
  const requests = [];

  contentArrays.forEach(function(contentArray) {
    traverseContent(contentArray, function(run) {
      const fgColor =
        run.style.foregroundColor &&
        run.style.foregroundColor.color &&
        run.style.foregroundColor.color.rgbColor;

      if (!fgColor) return;

      colorMap.forEach(function(mapping) {
        if (normalizedRgbMatches(fgColor, mapping.oldHex)) {
          requests.push({
            updateTextStyle: {
              range: { startIndex: run.startIndex, endIndex: run.endIndex },
              textStyle: {
                foregroundColor: {
                  color: { rgbColor: hexToNormalizedRgb(mapping.newHex) },
                },
              },
              fields: "foregroundColor",
            },
          });
        }
      });
    });
  });

  return requests;
}

// ---------------------------------------------------------------------------
// Step 4 — buildDocFontRequests
// ---------------------------------------------------------------------------

/**
 * Iterates all content sources and builds updateTextStyle request objects for
 * every textRun whose explicit font family matches an entry in fontMap.
 * Runs with a null font family (inheriting from Named Style) are skipped.
 * Font weight is preserved via weightedFontFamily when present.
 *
 * @param {Object[]} contentArrays  Array of .content arrays to traverse.
 * @param {Object[]} fontMap        Array of { oldFont, newFont } entries (FONT_MAP).
 * @returns {Object[]}              Array of updateTextStyle request objects.
 */
function buildDocFontRequests(contentArrays, fontMap) {
  const requests = [];

  contentArrays.forEach(function(contentArray) {
    traverseContent(contentArray, function(run) {
      const style = run.style;

      // Prefer weightedFontFamily (includes weight), fall back to fontFamily
      const wff = style.weightedFontFamily;
      const fontFamily = wff ? wff.fontFamily : style.fontFamily;
      if (!fontFamily) return; // null means inheriting — leave untouched

      fontMap.forEach(function(mapping) {
        if (fontFamily === mapping.oldFont) {
          const existingWeight = wff ? wff.weight : 400;
          requests.push({
            updateTextStyle: {
              range: { startIndex: run.startIndex, endIndex: run.endIndex },
              textStyle: {
                weightedFontFamily: {
                  fontFamily: mapping.newFont,
                  weight: existingWeight,
                },
              },
              fields: "weightedFontFamily",
            },
          });
        }
      });
    });
  });

  return requests;
}

// ---------------------------------------------------------------------------
// Step 5 — buildDocLogoRequests
// ---------------------------------------------------------------------------

/**
 * Searches doc.inlineObjects for embedded images whose sourceUri matches the
 * old logo URL pattern, and returns updateInlineObjectProperties requests that
 * patch sourceUri to the new logo Drive URL.
 *
 * NOTE: updateInlineObjectProperties.sourceUri updates only the link metadata,
 * not the embedded pixel data displayed in the document. Full pixel re-embedding
 * (delete + re-insert) is a stretch goal for a later phase.
 *
 * @param {Object} doc     Full document object from Docs.Documents.get().
 * @param {Object} config  LOGO_CONFIG from utils.js (uses newLogoFileId).
 * @returns {Object[]}     Array of updateInlineObjectProperties request objects.
 */
function buildDocLogoRequests(doc, config) {
  const requests = [];
  const inlineObjects = doc.inlineObjects;
  if (!inlineObjects) return requests;

  const newLogoUrl = driveFileUrl(config.newLogoFileId);
  // Match any drive.google.com URL that does not already point to the new file
  const OLD_LOGO_PATTERN = /drive\.google\.com/i;

  Object.keys(inlineObjects).forEach(function(objectId) {
    const embeddedObject =
      inlineObjects[objectId].inlineObjectProperties &&
      inlineObjects[objectId].inlineObjectProperties.embeddedObject;

    if (!embeddedObject) return;

    const sourceUri =
      (embeddedObject.imageProperties && embeddedObject.imageProperties.sourceUri) ||
      embeddedObject.sourceUri;

    if (!sourceUri) return;

    // Only replace if it looks like a Drive-hosted image and is not already the new logo
    if (OLD_LOGO_PATTERN.test(sourceUri) && sourceUri.indexOf(config.newLogoFileId) === -1) {
      requests.push({
        updateInlineObjectProperties: {
          objectId: objectId,
          inlineObjectProperties: {
            embeddedObject: {
              imageProperties: { sourceUri: newLogoUrl },
            },
          },
          fields: "embeddedObject.imageProperties.sourceUri",
        },
      });
    }
  });

  return requests;
}

// ---------------------------------------------------------------------------
// Step 6 — updateDocsDocument (public orchestrator)
// ---------------------------------------------------------------------------

/**
 * Runs the full brand update pipeline on a single Google Doc:
 *   1. Fetch document via Docs API (includeTabsContent: true)
 *   2. Collect content arrays (body + headers + footers)
 *   3. Build color, font, and logo requests
 *   4. batchUpdate the document
 *
 * @param {string} documentId
 */
function updateDocsDocument(documentId) {
  Logger.log("Starting brand update for document: %s", documentId);

  const doc = Docs.Documents.get(documentId, { includeTabsContent: true });

  // Collect all content arrays: body + headers + footers
  const contentArrays = [];

  if (doc.body && doc.body.content) {
    contentArrays.push(doc.body.content);
  }

  const headers = doc.headers || {};
  Object.keys(headers).forEach(function(key) {
    if (headers[key].content) contentArrays.push(headers[key].content);
  });

  const footers = doc.footers || {};
  Object.keys(footers).forEach(function(key) {
    if (footers[key].content) contentArrays.push(footers[key].content);
  });

  const colorReqs = buildDocColorRequests(contentArrays, COLOR_MAP);
  const fontReqs  = buildDocFontRequests(contentArrays, FONT_MAP);
  const logoReqs  = buildDocLogoRequests(doc, LOGO_CONFIG);

  const allRequests = [].concat(colorReqs, fontReqs, logoReqs);

  if (allRequests.length === 0) {
    Logger.log("  No changes needed for document: %s", documentId);
    return;
  }

  Docs.Documents.batchUpdate({ requests: allRequests }, documentId);

  Logger.log(
    "  ✓ Document updated: %d color, %d font, %d logo requests",
    colorReqs.length,
    fontReqs.length,
    logoReqs.length
  );
}
