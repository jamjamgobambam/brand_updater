// =============================================================================
// docs-updater.js — Google Docs brand updater (colors, fonts, logos)
// Depends on globals defined in utils.js: COLOR_MAP, FONT_MAP, LOGO_CONFIG,
// normalizedRgbMatches, hexToNormalizedRgb, driveFileUrl
// Requires the Docs Advanced Service (userSymbol: "Docs") enabled in the
// Apps Script project and appsscript.json.
// =============================================================================

// ---------------------------------------------------------------------------
// Step 4 — collectDocContent
// ---------------------------------------------------------------------------

/**
 * Returns a flat array of { content, segmentId } pairs for all segments of
 * the document: body, headers, footers, and footnotes.
 *
 * segmentId is "" for the body; the header / footer / footnote's own opaque
 * ID for all other segments. Every range-based Docs API request requires the
 * correct segmentId — omitting it causes a 400 error for non-body segments.
 *
 * @param {Object} document  Full document object from Docs.Documents.get().
 * @returns {{ content: Object[], segmentId: string }[]}
 */
function collectDocContent(document) {
  const segments = [];

  if (document.body && document.body.content) {
    segments.push({ content: document.body.content, segmentId: "" });
  }

  const headers = document.headers || {};
  Object.keys(headers).forEach(function(headerId) {
    if (headers[headerId].content) {
      segments.push({ content: headers[headerId].content, segmentId: headerId });
    }
  });

  const footers = document.footers || {};
  Object.keys(footers).forEach(function(footerId) {
    if (footers[footerId].content) {
      segments.push({ content: footers[footerId].content, segmentId: footerId });
    }
  });

  const footnotes = document.footnotes || {};
  Object.keys(footnotes).forEach(function(footnoteId) {
    if (footnotes[footnoteId].content) {
      segments.push({ content: footnotes[footnoteId].content, segmentId: footnoteId });
    }
  });

  return segments;
}

// ---------------------------------------------------------------------------
// Internal traversal helper
// ---------------------------------------------------------------------------

/**
 * Walks a single Docs content array, calling callback for each textRun found
 * in paragraphs and table cells (recursively handles nested tables).
 *
 * @param {Object[]} contentArray
 * @param {Function} callback  Called with { startIndex, endIndex, style }.
 */
function traverseContentArray(contentArray, callback) {
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
          traverseContentArray(cell.content, callback);
        });
      });
    }
  });
}

// ---------------------------------------------------------------------------
// Step 5 — buildDocColorRequests
// ---------------------------------------------------------------------------

/**
 * Builds updateTextStyle requests for every textRun whose explicit foreground
 * color matches an entry in colorMap. segmentId is included in every range so
 * requests targeting headers, footers, and footnotes are accepted by the API.
 *
 * Only explicit inline foregroundColor overrides are changed; text whose color
 * is inherited from a Named Style has no rgbColor here and is not touched.
 *
 * @param {{ content: Object[], segmentId: string }[]} segments
 * @param {Object[]} colorMap  Array of { oldHex, newHex } entries (COLOR_MAP).
 * @returns {Object[]}         Array of updateTextStyle request objects.
 */
function buildDocColorRequests(segments, colorMap) {
  const requests = [];

  segments.forEach(function(segment) {
    traverseContentArray(segment.content, function(run) {
      const fgColor =
        run.style.foregroundColor &&
        run.style.foregroundColor.color &&
        run.style.foregroundColor.color.rgbColor;

      if (!fgColor) return;

      colorMap.forEach(function(mapping) {
        if (normalizedRgbMatches(fgColor, mapping.oldHex)) {
          requests.push({
            updateTextStyle: {
              range: {
                startIndex: run.startIndex,
                endIndex:   run.endIndex,
                segmentId:  segment.segmentId,
              },
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
// Step 6 — replaceDocColors
// ---------------------------------------------------------------------------

/**
 * Fetches a document and submits all color replacement requests in a single
 * batchUpdate call.
 *
 * @param {string} docId
 */
function replaceDocColors(docId) {
  const doc      = Docs.Documents.get(docId);
  const segments = collectDocContent(doc);
  const requests = buildDocColorRequests(segments, COLOR_MAP);

  if (requests.length === 0) {
    Logger.log("  replaceDocColors: no color changes for %s", docId);
    return;
  }

  Docs.Documents.batchUpdate({ requests: requests }, docId);
  Logger.log("  replaceDocColors: %d requests submitted for %s", requests.length, docId);
}

// ---------------------------------------------------------------------------
// Step 4 — buildDocFontRequests
// ---------------------------------------------------------------------------
// Step 7 — buildDocFontRequests
// ---------------------------------------------------------------------------

/**
 * Builds updateTextStyle requests for every textRun whose explicit font family
 * matches an entry in fontMap. Preserves font weight via weightedFontFamily so
 * bold runs stay bold after replacement.
 *
 * Runs with no explicit font (inheriting from a Named Style) are skipped.
 * segmentId is included in every range.
 *
 * @param {{ content: Object[], segmentId: string }[]} segments
 * @param {Object[]} fontMap  Array of { oldFont, newFont } entries (FONT_MAP).
 * @returns {Object[]}        Array of updateTextStyle request objects.
 */
function buildDocFontRequests(segments, fontMap) {
  const requests = [];

  segments.forEach(function(segment) {
    traverseContentArray(segment.content, function(run) {
      const style = run.style;
      const wff   = style.weightedFontFamily;
      const fontFamily = wff ? wff.fontFamily : style.fontFamily;
      if (!fontFamily) return; // inheriting from Named Style — skip

      fontMap.forEach(function(mapping) {
        if (fontFamily === mapping.oldFont) {
          const existingWeight = wff ? wff.weight : 400;
          requests.push({
            updateTextStyle: {
              range: {
                startIndex: run.startIndex,
                endIndex:   run.endIndex,
                segmentId:  segment.segmentId,
              },
              textStyle: {
                weightedFontFamily: {
                  fontFamily: mapping.newFont,
                  weight:     existingWeight,
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
// Step 8 — replaceDocFonts
// ---------------------------------------------------------------------------

/**
 * Fetches a document and submits all font replacement requests in a single
 * batchUpdate call.
 *
 * @param {string} docId
 */
function replaceDocFonts(docId) {
  const doc      = Docs.Documents.get(docId);
  const segments = collectDocContent(doc);
  const requests = buildDocFontRequests(segments, FONT_MAP);

  if (requests.length === 0) {
    Logger.log("  replaceDocFonts: no font changes for %s", docId);
    return;
  }

  Docs.Documents.batchUpdate({ requests: requests }, docId);
  Logger.log("  replaceDocFonts: %d requests submitted for %s", requests.length, docId);
}

// ---------------------------------------------------------------------------
// Step 9 — logDocImages
// ---------------------------------------------------------------------------

/**
 * Diagnostic utility — run once on a representative document to discover
 * sourceUri values and dimensions of all inline images, then use that data
 * to configure LOGO_CONFIG.docsLogo in utils.js.
 *
 * Logs: objectId, segment (body / header / footer), sourceUri, width (PT),
 * and height (PT) for every inline image in the document.
 * Makes no changes to the document.
 *
 * @param {string} docId
 */
function logDocImages(docId) {
  const doc           = Docs.Documents.get(docId);
  const inlineObjects = doc.inlineObjects || {};

  // Build a map of objectId → segment label for reporting
  const objectSegment = {};

  function scanForImageElements(contentArray, segmentLabel) {
    if (!contentArray) return;
    contentArray.forEach(function(el) {
      if (el.paragraph) {
        (el.paragraph.elements || []).forEach(function(pe) {
          if (pe.inlineObjectElement) {
            objectSegment[pe.inlineObjectElement.inlineObjectId] = segmentLabel;
          }
        });
      }
      if (el.table) {
        (el.table.tableRows || []).forEach(function(row) {
          (row.tableCells || []).forEach(function(cell) {
            scanForImageElements(cell.content, segmentLabel);
          });
        });
      }
    });
  }

  if (doc.body) scanForImageElements(doc.body.content, "body");
  Object.keys(doc.headers || {}).forEach(function(id) {
    scanForImageElements(doc.headers[id].content, "header:" + id);
  });
  Object.keys(doc.footers || {}).forEach(function(id) {
    scanForImageElements(doc.footers[id].content, "footer:" + id);
  });

  Object.keys(inlineObjects).forEach(function(objectId) {
    const embedded =
      inlineObjects[objectId].inlineObjectProperties &&
      inlineObjects[objectId].inlineObjectProperties.embeddedObject;
    if (!embedded) return;

    const sourceUri = embedded.imageProperties && embedded.imageProperties.sourceUri;
    const width     = embedded.size && embedded.size.width  && embedded.size.width.magnitude;
    const height    = embedded.size && embedded.size.height && embedded.size.height.magnitude;
    const unit      = embedded.size && embedded.size.width  && embedded.size.width.unit;
    const segment   = objectSegment[objectId] || "unknown";

    Logger.log(
      "Image — objectId: %s | segment: %s | sourceUri: %s | width: %s %s | height: %s %s",
      objectId, segment, sourceUri || "(null)", width, unit, height, unit
    );
  });
}

// ---------------------------------------------------------------------------
// Step 11 — buildDocLogoRequests
// ---------------------------------------------------------------------------

/**
 * Finds all logo inline objects in body + headers + footers, and returns
 * deleteContentRange + insertInlineImage request pairs sorted in reverse
 * startIndex order to prevent index-shift bugs during batchUpdate.
 *
 * Matching (checked in order):
 *   Primary:  sourceUri === LOGO_CONFIG.docsLogo.oldSourceUri (when non-null)
 *   Fallback: width and height within configured PT bounds
 *
 * @param {Object}  doc         Full document from Docs.Documents.get().
 * @param {string}  newLogoUrl  Drive export URL for the replacement logo.
 * @param {boolean} dryRun      If true, log matches but return no requests.
 * @returns {Object[]}          Flat array of request objects, reverse-index order.
 */
function buildDocLogoRequests(doc, newLogoUrl, dryRun) {
  const logoConfig    = LOGO_CONFIG.docsLogo;
  const inlineObjects = doc.inlineObjects || {};

  // Step 1 — build reverse index: position of every inlineObjectElement
  const reverseIndex = [];

  function indexContentForImages(contentArray, segmentId) {
    if (!contentArray) return;
    contentArray.forEach(function(el) {
      if (el.paragraph) {
        (el.paragraph.elements || []).forEach(function(pe) {
          if (pe.inlineObjectElement) {
            reverseIndex.push({
              objectId:   pe.inlineObjectElement.inlineObjectId,
              startIndex: pe.startIndex,
              segmentId:  segmentId,
            });
          }
        });
      }
      if (el.table) {
        (el.table.tableRows || []).forEach(function(row) {
          (row.tableCells || []).forEach(function(cell) {
            indexContentForImages(cell.content, segmentId);
          });
        });
      }
    });
  }

  if (doc.body) indexContentForImages(doc.body.content, "");
  Object.keys(doc.headers || {}).forEach(function(id) {
    indexContentForImages(doc.headers[id].content, id);
  });
  Object.keys(doc.footers || {}).forEach(function(id) {
    indexContentForImages(doc.footers[id].content, id);
  });

  // Steps 2–4 — match inline objects and collect logo positions
  const matches = [];

  reverseIndex.forEach(function(entry) {
    const inlineObj = inlineObjects[entry.objectId];
    if (!inlineObj) return;

    const embedded =
      inlineObj.inlineObjectProperties &&
      inlineObj.inlineObjectProperties.embeddedObject;
    if (!embedded) return;

    const sourceUri = embedded.imageProperties && embedded.imageProperties.sourceUri;
    const widthPt   = embedded.size && embedded.size.width  && embedded.size.width.magnitude;
    const heightPt  = embedded.size && embedded.size.height && embedded.size.height.magnitude;

    let isMatch = false;
    if (logoConfig.oldSourceUri !== null && logoConfig.oldSourceUri !== undefined) {
      isMatch = (sourceUri === logoConfig.oldSourceUri);
    } else {
      isMatch =
        widthPt  >= logoConfig.minWidthPt  &&
        widthPt  <= logoConfig.maxWidthPt  &&
        heightPt >= logoConfig.minHeightPt &&
        heightPt <= logoConfig.maxHeightPt;
    }

    if (!isMatch) return;

    if (dryRun) {
      Logger.log(
        "DRY RUN — logo match: objectId=%s segmentId=%s startIndex=%s sourceUri=%s width=%sPT height=%sPT",
        entry.objectId, entry.segmentId, entry.startIndex,
        sourceUri || "(null)", widthPt, heightPt
      );
      return;
    }

    matches.push({
      startIndex: entry.startIndex,
      segmentId:  entry.segmentId,
      widthPt:    widthPt,
      heightPt:   heightPt,
    });
  });

  if (dryRun) return [];

  // Step 5 — sort in reverse order to prevent index shifts from invalidating later operations
  matches.sort(function(a, b) { return b.startIndex - a.startIndex; });

  // Build flat request array: delete then insert for each match
  const requests = [];
  matches.forEach(function(match) {
    requests.push({
      deleteContentRange: {
        range: {
          startIndex: match.startIndex,
          endIndex:   match.startIndex + 1,
          segmentId:  match.segmentId,
        },
      },
    });
    requests.push({
      insertInlineImage: {
        location: {
          index:     match.startIndex,
          segmentId: match.segmentId,
        },
        uri: newLogoUrl,
        objectSize: {
          width:  { magnitude: match.widthPt,  unit: "PT" },
          height: { magnitude: match.heightPt, unit: "PT" },
        },
      },
    });
  });

  return requests;
}

// ---------------------------------------------------------------------------
// Step 12 — replaceDocLogos
// ---------------------------------------------------------------------------

/**
 * Fetches a document and submits logo delete + insert requests in a single
 * batchUpdate call.
 *
 * @param {string}  docId
 * @param {boolean} [dryRun=false]  If true, logs matches but makes no changes.
 */
function replaceDocLogos(docId, dryRun) {
  const doc        = Docs.Documents.get(docId);
  const newLogoUrl = driveFileUrl(LOGO_CONFIG.newLogoFileId);
  const requests   = buildDocLogoRequests(doc, newLogoUrl, dryRun);

  if (dryRun) {
    Logger.log("  replaceDocLogos: dry run complete for %s", docId);
    return;
  }

  if (requests.length === 0) {
    Logger.log("  replaceDocLogos: no logo matches for %s", docId);
    return;
  }

  Docs.Documents.batchUpdate({ requests: requests }, docId);
  Logger.log("  replaceDocLogos: %d requests submitted for %s", requests.length, docId);
}

// ---------------------------------------------------------------------------
// Step 13 — updateDocsDocument (public orchestrator)
// ---------------------------------------------------------------------------

/**
 * Runs the full brand update pipeline on a single Google Doc:
 *   1. replaceDocColors  — explicit inline foreground color overrides
 *   2. replaceDocFonts   — Poppins / Figtree → Lexend
 *   3. replaceDocLogos   — delete + re-insert logo images
 *
 * Each step can be called independently for isolated testing.
 * The dryRun flag is passed through to replaceDocLogos only.
 *
 * @param {string}  docId
 * @param {boolean} [dryRun=false]  Passed through to replaceDocLogos.
 */
function updateDocsDocument(docId, dryRun) {
  Logger.log("Starting brand update for document: %s", docId);
  replaceDocColors(docId);
  replaceDocFonts(docId);
  replaceDocLogos(docId, dryRun);
  Logger.log("Done: %s", docId);
}

