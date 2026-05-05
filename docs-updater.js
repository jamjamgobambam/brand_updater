// =============================================================================
// docs-updater.js — Google Docs brand updater (colors, fonts, logos)
// Depends on globals defined in utils.js: COLOR_MAP, FONT_MAP, LOGO_CONFIG,
// normalizedRgbMatches, hexToNormalizedRgb
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
// buildNamedStyleLookup
// ---------------------------------------------------------------------------

/**
 * Builds a lookup map of namedStyleType → textStyle from the document's
 * namedStyles array. Used as a fallback when a text run carries no explicit
 * inline style, so the effective font / color can still be detected.
 *
 * Keys are namedStyleType strings (e.g. "NORMAL_TEXT", "HEADING_1").
 * Values are the textStyle object for that style (may be {}).
 *
 * @param {Object} document  Full document object from Docs.Documents.get().
 * @returns {Object}  Map: namedStyleType → textStyle object.
 */
function buildNamedStyleLookup(document) {
  const lookup = {};
  const styles = (document.namedStyles && document.namedStyles.styles) || [];
  styles.forEach(function(namedStyle) {
    if (namedStyle.namedStyleType) {
      lookup[namedStyle.namedStyleType] = namedStyle.textStyle || {};
    }
  });
  return lookup;
}

// ---------------------------------------------------------------------------
// Internal traversal helper
// ---------------------------------------------------------------------------

/**
 * Walks a single Docs content array, calling callback for each textRun found
 * in paragraphs and table cells (recursively handles nested tables).
 *
 * @param {Object[]} contentArray
 * @param {Function} callback  Called with { startIndex, endIndex, style, namedStyleType }.
 */
function traverseContentArray(contentArray, callback) {
  if (!contentArray) return;

  contentArray.forEach(function(structuralElement) {
    if (structuralElement.paragraph) {
      var namedStyleType =
        (structuralElement.paragraph.paragraphStyle &&
         structuralElement.paragraph.paragraphStyle.namedStyleType) ||
        "NORMAL_TEXT";
      (structuralElement.paragraph.elements || []).forEach(function(element) {
        if (!element.textRun) return;
        // Google Docs API omits startIndex when it is 0 (default value elision).
        // Coerce undefined → 0 so batchUpdate range requests are always valid.
        var startIndex = element.startIndex !== undefined ? element.startIndex : 0;
        callback({
          startIndex:     startIndex,
          endIndex:       element.endIndex,
          style:          element.textRun.textStyle || {},
          namedStyleType: namedStyleType,
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

/**
 * Walks all structural elements in contentArray, calling onElement(el, segmentId)
 * for each, then recursively descending into any table's cell content.
 * Used as a shared traversal backbone for per-element logic that is not specific
 * to textRuns (e.g. paragraph shading, table cell styles, inline images).
 *
 * @param {Object[]} contentArray  Array of structural elements (may be null/undefined).
 * @param {string}   segmentId     Docs API segment ID ("" for body, opaque for others).
 * @param {Function} onElement     Callback invoked as onElement(el, segmentId) for each element.
 */
function walkDocContent(contentArray, segmentId, onElement) {
  if (!contentArray) return;
  contentArray.forEach(function(el) {
    onElement(el, segmentId);
    if (el.table) {
      (el.table.tableRows || []).forEach(function(row) {
        (row.tableCells || []).forEach(function(cell) {
          walkDocContent(cell.content, segmentId, onElement);
        });
      });
    }
  });
}

// ---------------------------------------------------------------------------
// batchUpdateDocWithUrlFetch
// ---------------------------------------------------------------------------

/**
 * Sends a batchUpdate to the Docs REST API directly via UrlFetchApp, bypassing
 * the Apps Script Advanced Service wrapper.
 *
 * The Advanced Service converts camelCase JS keys and may not support all
 * request types (e.g. updateNamedStyle). UrlFetchApp sends the JSON payload
 * exactly as built, preserving camelCase field names as the REST API expects.
 *
 * Requires that the script has already been granted documents scope (satisfied
 * automatically when any Docs Advanced Service call has been made).
 *
 * @param {string}   docId     Google Docs document ID.
 * @param {Object[]} requests  Array of Docs API request objects (camelCase).
 */
function batchUpdateDocWithUrlFetch(docId, requests) {
  // Drop any null / undefined / empty-object entries that would cause a 400.
  var clean = requests.filter(function(r) { return r && typeof r === "object" && Object.keys(r).length > 0; });
  if (clean.length === 0) return;
  var token   = ScriptApp.getOAuthToken();
  var url     = "https://docs.googleapis.com/v1/documents/" + docId + ":batchUpdate";
  var payload = JSON.stringify({ requests: clean });
  var response = UrlFetchApp.fetch(url, {
    method:             "post",
    contentType:        "application/json",
    headers:            { Authorization: "Bearer " + token },
    payload:            payload,
    muteHttpExceptions: true,
  });
  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error("batchUpdate (REST) failed (" + code + "): " + response.getContentText());
  }
}

// ---------------------------------------------------------------------------
// buildDocTableCellColorRequests
// ---------------------------------------------------------------------------

/**
 * Walks all tables in all segments and builds updateTableCellStyle requests
 * for every cell whose background color or border colors match an entry in
 * colorMap.
 *
 * Must be sent via batchUpdateDocWithUrlFetch (REST) because the Apps Script
 * Advanced Service does not support updateTableCellStyle.
 *
 * @param {Object}   doc       Full document from Docs.Documents.get().
 * @param {Object[]} colorMap  Array of { oldHex, newHex } entries.
 * @returns {Object[]}         Array of updateTableCellStyle request objects.
 */
function buildDocTableCellColorRequests(doc, colorMap) {
  const requests = [];
  const BORDER_SIDES = ["borderLeft", "borderRight", "borderTop", "borderBottom"];

  function getBorderRgb(cell, side) {
    return (
      cell.tableCellStyle &&
      cell.tableCellStyle[side] &&
      cell.tableCellStyle[side].color &&
      cell.tableCellStyle[side].color.color &&
      cell.tableCellStyle[side].color.color.rgbColor
    ) || null;
  }

  function processTableCells(el, segmentId) {
    if (!el.table) return;
    (el.table.tableRows || []).forEach(function(row, rowIndex) {
      (row.tableCells || []).forEach(function(cell, colIndex) {
        const tableStartIndex = el.startIndex !== undefined ? el.startIndex : 0;
        const tableRange = {
          tableCellLocation: {
            tableStartLocation: { index: tableStartIndex, segmentId: segmentId },
            rowIndex:           rowIndex,
            columnIndex:        colIndex,
          },
          rowSpan:    1,
          columnSpan: 1,
        };

        // --- Background color ---
        const bgColor =
          cell.tableCellStyle &&
          cell.tableCellStyle.backgroundColor &&
          cell.tableCellStyle.backgroundColor.color &&
          cell.tableCellStyle.backgroundColor.color.rgbColor;

        if (bgColor) {
          var bgNewHex = findColorMapping(bgColor, colorMap, COLOR_DISTANCE_THRESHOLD);
          if (bgNewHex) {
            requests.push({
              updateTableCellStyle: {
                tableRange:     tableRange,
                tableCellStyle: {
                  backgroundColor: { color: { rgbColor: hexToNormalizedRgb(bgNewHex) } },
                },
                fields: "backgroundColor",
              },
            });
          }
        }

        // --- Border colors (one request per side that has a matching color) ---
        BORDER_SIDES.forEach(function(side) {
          const borderRgb = getBorderRgb(cell, side);
          if (!borderRgb) return;

          var borderNewHex = findColorMapping(borderRgb, colorMap, COLOR_DISTANCE_THRESHOLD);
          if (borderNewHex) {
            const newBorder = Object.assign(
              {},
              cell.tableCellStyle[side],
              { color: { color: { rgbColor: hexToNormalizedRgb(borderNewHex) } } }
            );
            const stylePatch = {};
            stylePatch[side] = newBorder;
            requests.push({
              updateTableCellStyle: {
                tableRange:     tableRange,
                tableCellStyle: stylePatch,
                fields:         side,
              },
            });
          }
        });
      });
    });
  }

  walkDocContent(doc.body ? doc.body.content : null, "", processTableCells);
  Object.keys(doc.headers || {}).forEach(function(id) {
    walkDocContent(doc.headers[id].content, id, processTableCells);
  });
  Object.keys(doc.footers || {}).forEach(function(id) {
    walkDocContent(doc.footers[id].content, id, processTableCells);
  });

  return requests;
}

// ---------------------------------------------------------------------------
// Step 5 — buildDocColorRequests
// ---------------------------------------------------------------------------

/**
 * Builds updateTextStyle requests for every textRun whose effective
 * foreground color matches an entry in colorMap.
 *
 * Three-level probe per run:
 *   1. Explicit inline foregroundColor override on the run.
 *   2. foregroundColor in the Named Style for the paragraph's namedStyleType.
 *      (If the Named Style color is not in colorMap, fall through to level 3.)
 *   3. NORMAL_TEXT Named Style color as a proxy for theme-inherited values.
 *      This catches TITLE and HEADING paragraphs that carry no inline or
 *      Named-Style color but display the document's default brand color.
 *
 * @param {{ content: Object[], segmentId: string }[]} segments
 * @param {Object[]} colorMap        Array of { oldHex, newHex } entries.
 * @param {Object}   namedStyleLookup  Map of namedStyleType → textStyle.
 * @returns {Object[]}  Array of updateTextStyle request objects.
 */
function buildDocColorRequests(segments, colorMap, namedStyleLookup) {
  const requests = [];
  const normalStyle = (namedStyleLookup && namedStyleLookup["NORMAL_TEXT"]) || {};
  const normalRgb   =
    normalStyle.foregroundColor &&
    normalStyle.foregroundColor.color &&
    normalStyle.foregroundColor.color.rgbColor;

  segments.forEach(function(segment) {
    traverseContentArray(segment.content, function(run) {
      // Level 1: explicit inline color.
      const explicitRgb =
        run.style.foregroundColor &&
        run.style.foregroundColor.color &&
        run.style.foregroundColor.color.rgbColor;

      var effectiveRgb = explicitRgb;

      if (!effectiveRgb && namedStyleLookup) {
        // Level 2: Named Style for this paragraph type.
        const nsStyle = namedStyleLookup[run.namedStyleType] || {};
        const nsRgb   =
          nsStyle.foregroundColor &&
          nsStyle.foregroundColor.color &&
          nsStyle.foregroundColor.color.rgbColor;
        if (nsRgb && findColorMapping(nsRgb, colorMap, COLOR_DISTANCE_THRESHOLD) !== null) {
          effectiveRgb = nsRgb;
        } else {
          // Level 3: NORMAL_TEXT proxy (catches theme-inherited heading colors
          // and TITLE whose Named Style has a non-brand explicit color).
          effectiveRgb = normalRgb;
        }
      }

      if (!effectiveRgb) return;

      var fgNewHex = findColorMapping(effectiveRgb, colorMap, COLOR_DISTANCE_THRESHOLD);
      if (fgNewHex) {
        requests.push({
          updateTextStyle: {
            range: {
              startIndex: run.startIndex,
              endIndex:   run.endIndex,
              segmentId:  segment.segmentId,
            },
            textStyle: {
              foregroundColor: {
                color: { rgbColor: hexToNormalizedRgb(fgNewHex) },
              },
            },
            fields: "foregroundColor",
          },
        });
      }

      // Text highlight (textStyle.backgroundColor)
      const highlightRgb =
        run.style.backgroundColor &&
        run.style.backgroundColor.color &&
        run.style.backgroundColor.color.rgbColor;
      if (highlightRgb) {
        var hlNewHex = findColorMapping(highlightRgb, colorMap, COLOR_DISTANCE_THRESHOLD);
        if (hlNewHex) {
          requests.push({
            updateTextStyle: {
              range: {
                startIndex: run.startIndex,
                endIndex:   run.endIndex,
                segmentId:  segment.segmentId,
              },
              textStyle: {
                backgroundColor: {
                  color: { rgbColor: hexToNormalizedRgb(hlNewHex) },
                },
              },
              fields: "backgroundColor",
            },
          });
        }
      }
    });
  });

  return requests;
}

// ---------------------------------------------------------------------------
// buildDocNamedStyleColorRequests
// ---------------------------------------------------------------------------

/**
 * Builds updateNamedStyle requests for every Named Style entry in the document
 * whose effective foreground color matches an entry in colorMap.
 *
 * For Named Styles that carry no explicit foregroundColor (i.e. the color is
 * theme-inherited), falls back to the NORMAL_TEXT Named Style's color as a
 * proxy for the document default. This covers heading styles (HEADING_1–6,
 * TITLE, SUBTITLE) that derive their color from the theme rather than storing
 * an explicit override in the named style definition.
 *
 * Updating the Named Style definition is the correct API approach for
 * restyling headings and titles, because their text runs carry no explicit
 * inline overrides and cannot be targeted by updateTextStyle.
 *
 * @param {Object}   doc       Full document object from Docs.Documents.get().
 * @param {Object[]} colorMap  Array of { oldHex, newHex } entries (COLOR_MAP).
 * @returns {Object[]}         Array of updateNamedStyle request objects.
 */
function buildDocNamedStyleColorRequests(doc, colorMap) {
  const requests = [];
  const styles   = (doc.namedStyles && doc.namedStyles.styles) || [];

  // Determine NORMAL_TEXT color as fallback for styles with no explicit color.
  var normalTextColor = null;
  styles.forEach(function(ns) {
    if (ns.namedStyleType === "NORMAL_TEXT") {
      normalTextColor =
        ns.textStyle &&
        ns.textStyle.foregroundColor &&
        ns.textStyle.foregroundColor.color &&
        ns.textStyle.foregroundColor.color.rgbColor;
    }
  });

  styles.forEach(function(ns) {
    const explicitColor =
      ns.textStyle &&
      ns.textStyle.foregroundColor &&
      ns.textStyle.foregroundColor.color &&
      ns.textStyle.foregroundColor.color.rgbColor;

    // For non-NORMAL_TEXT styles with no explicit color, proxy against
    // NORMAL_TEXT so theme-inherited heading colors are still detected.
    const effectiveColor = explicitColor ||
      (ns.namedStyleType !== "NORMAL_TEXT" ? normalTextColor : null);

    if (!effectiveColor) return;

    var nsColorNewHex = findColorMapping(effectiveColor, colorMap, COLOR_DISTANCE_THRESHOLD);
    if (nsColorNewHex) {
      requests.push({
        updateNamedStyle: {
          namedStyle: {
            namedStyleType: ns.namedStyleType,
            textStyle: {
              foregroundColor: {
                color: { rgbColor: hexToNormalizedRgb(nsColorNewHex) },
              },
            },
          },
          fields: "textStyle.foregroundColor",
        },
      });
    }
  });

  return requests;
}

// ---------------------------------------------------------------------------
// Step 6 — replaceDocColors
// ---------------------------------------------------------------------------

/**
 * Builds updateParagraphStyle requests for every paragraph whose shading
 * background color matches an entry in colorMap.
 *
 * @param {{ content: Object[], segmentId: string }[]} segments
 * @param {Object[]} colorMap  Array of { oldHex, newHex } entries.
 * @returns {Object[]}
 */
function buildDocParagraphShadingRequests(segments, colorMap) {
  var requests = [];

  segments.forEach(function(segment) {
    walkDocContent(segment.content, segment.segmentId, function(el, segmentId) {
      if (!el.paragraph) return;
      var shading =
        el.paragraph.paragraphStyle &&
        el.paragraph.paragraphStyle.shading &&
        el.paragraph.paragraphStyle.shading.backgroundColor &&
        el.paragraph.paragraphStyle.shading.backgroundColor.color &&
        el.paragraph.paragraphStyle.shading.backgroundColor.color.rgbColor;
      if (shading) {
        var shadingNewHex = findColorMapping(shading, colorMap, COLOR_DISTANCE_THRESHOLD);
        if (shadingNewHex) {
          var startIndex = el.startIndex !== undefined ? el.startIndex : 0;
          requests.push({
            updateParagraphStyle: {
              range: {
                startIndex: startIndex,
                endIndex:   el.endIndex,
                segmentId:  segmentId,
              },
              paragraphStyle: {
                shading: {
                  backgroundColor: {
                    color: { rgbColor: hexToNormalizedRgb(shadingNewHex) },
                  },
                },
              },
              fields: "shading.backgroundColor",
            },
          });
        }
      }
    });
  });

  return requests;
}

/**
 * Builds an updateDocumentStyle request if the page background color matches
 * an entry in colorMap. Returns an array of 0 or 1 requests.
 *
 * @param {Object}   doc       Full document object from Docs.Documents.get().
 * @param {Object[]} colorMap  Array of { oldHex, newHex } entries.
 * @returns {Object[]}
 */
function buildDocPageBackgroundRequest(doc, colorMap) {
  var bgRgb =
    doc.documentStyle &&
    doc.documentStyle.background &&
    doc.documentStyle.background.color &&
    doc.documentStyle.background.color.rgbColor;

  if (!bgRgb) return [];

  var bgNewHex = findColorMapping(bgRgb, colorMap, COLOR_DISTANCE_THRESHOLD);
  if (!bgNewHex) return [];
  return [{
    updateDocumentStyle: {
      documentStyle: {
        background: {
          color: { rgbColor: hexToNormalizedRgb(bgNewHex) },
        },
      },
      fields: "background.color",
    },
  }];
}

/**
 * Fetches a document and submits all color replacement requests in a single
 * batchUpdate call.
 *
 * @param {string} docId
 */
function replaceDocColors(docId) {
  const doc      = Docs.Documents.get(docId);
  const segments = collectDocContent(doc);

  const nsLookup       = buildNamedStyleLookup(doc);
  const inlineReqs     = buildDocColorRequests(segments, COLOR_MAP, nsLookup);
  const cellReqs       = buildDocTableCellColorRequests(doc, COLOR_MAP);
  const shadingReqs    = buildDocParagraphShadingRequests(segments, COLOR_MAP);
  const pageBgReqs     = buildDocPageBackgroundRequest(doc, COLOR_MAP);

  if (inlineReqs.length === 0 && cellReqs.length === 0 &&
      shadingReqs.length === 0 && pageBgReqs.length === 0) {
    Logger.log("  replaceDocColors: no color changes for %s", docId);
    return;
  }

  // Text run foreground + highlight + paragraph shading via Advanced Service.
  const advancedReqs = inlineReqs.concat(shadingReqs);
  if (advancedReqs.length > 0) {
    Docs.Documents.batchUpdate({ requests: advancedReqs }, docId);
    Logger.log("  replaceDocColors: %d text/shading requests submitted for %s", advancedReqs.length, docId);
  }

  // Table-cell and page-background requests via REST (Advanced Service
  // wrapper does not support updateTableCellStyle or updateDocumentStyle
  // reliably with nested fields).
  const restReqs = cellReqs.concat(pageBgReqs);
  if (restReqs.length > 0) {
    try {
      batchUpdateDocWithUrlFetch(docId, restReqs);
      Logger.log("  replaceDocColors: %d cell/page-bg requests submitted for %s", restReqs.length, docId);
    } catch (e) {
      Logger.log("  cell/page-bg request FAILED (%s). First request: %s", e.message, JSON.stringify(restReqs[0]));
      throw e;
    }
  }
}

// ---------------------------------------------------------------------------
// Step 4 — buildDocFontRequests
// ---------------------------------------------------------------------------
// Step 7 — buildDocFontRequests
// ---------------------------------------------------------------------------

/**
 * Builds updateTextStyle requests for every textRun whose effective font
 * matches an entry in fontMap. Preserves weight so bold runs stay bold.
 *
 * Three-level probe per run:
 *   1. Explicit inline weightedFontFamily / fontFamily override on the run.
 *   2. weightedFontFamily in the Named Style for the paragraph's namedStyleType.
 *      (If the Named Style font is not in fontMap, fall through to level 3.)
 *   3. NORMAL_TEXT Named Style font as a proxy for theme-inherited values.
 *      This catches TITLE and HEADING paragraphs that carry no inline or
 *      Named-Style font but display the document's default brand font.
 *
 * @param {{ content: Object[], segmentId: string }[]} segments
 * @param {Object[]} fontMap         Array of { oldFont, newFont } entries.
 * @param {Object}   namedStyleLookup  Map of namedStyleType → textStyle.
 * @returns {Object[]}  Array of updateTextStyle request objects.
 */
function buildDocFontRequests(segments, fontMap, namedStyleLookup) {
  const requests = [];
  const normalStyle  = (namedStyleLookup && namedStyleLookup["NORMAL_TEXT"]) || {};
  const normalWff    = normalStyle.weightedFontFamily;
  const normalFamily = normalWff ? normalWff.fontFamily : normalStyle.fontFamily;
  const normalWeight = normalWff ? normalWff.weight : 400;

  segments.forEach(function(segment) {
    traverseContentArray(segment.content, function(run) {
      // Level 1: explicit inline font.
      const style          = run.style;
      const wff            = style.weightedFontFamily;
      const explicitFamily = wff ? wff.fontFamily : style.fontFamily;
      const explicitWeight = wff ? wff.weight : null;

      var effectiveFamily = explicitFamily;
      var effectiveWeight = explicitWeight;

      if (!effectiveFamily && namedStyleLookup) {
        // Level 2: Named Style for this paragraph type.
        const nsStyle  = namedStyleLookup[run.namedStyleType] || {};
        const nsWff    = nsStyle.weightedFontFamily;
        const nsFamily = nsWff ? nsWff.fontFamily : nsStyle.fontFamily;
        if (nsFamily && (fontMap.some(function(m) { return nsFamily === m.oldFont; }) || BRAND_FONTS.indexOf(nsFamily) === -1)) {
          effectiveFamily = nsFamily;
          effectiveWeight = nsWff ? nsWff.weight : 400;
        } else {
          // Level 3: NORMAL_TEXT proxy (catches theme-inherited heading fonts
          // and TITLE whose Named Style has a non-brand explicit font).
          effectiveFamily = normalFamily;
          effectiveWeight = normalWeight;
        }
      }

      if (!effectiveFamily) return;

      var docFontMatched = false;
      fontMap.forEach(function(mapping) {
        if (effectiveFamily === mapping.oldFont) {
          docFontMatched = true;
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
                  weight:     effectiveWeight || 400,
                },
              },
              fields: "weightedFontFamily",
            },
          });
        }
      });

      // Replace any non-brand font not handled by FONT_MAP
      if (!docFontMatched && BRAND_FONTS.indexOf(effectiveFamily) === -1) {
        requests.push({
          updateTextStyle: {
            range: {
              startIndex: run.startIndex,
              endIndex:   run.endIndex,
              segmentId:  segment.segmentId,
            },
            textStyle: {
              weightedFontFamily: {
                fontFamily: FALLBACK_FONT,
                weight:     effectiveWeight || 400,
              },
            },
            fields: "weightedFontFamily",
          },
        });
      }
    });
  });

  return requests;
}

// ---------------------------------------------------------------------------
// buildDocNamedStyleFontRequests
// ---------------------------------------------------------------------------

/**
 * Builds updateNamedStyle requests for every Named Style entry in the document
 * whose effective font matches an entry in fontMap.
 *
 * For Named Styles that carry no explicit weightedFontFamily (i.e. the font is
 * theme-inherited), falls back to the NORMAL_TEXT Named Style's font as a
 * proxy for the document default. This covers heading styles (HEADING_1–6,
 * TITLE, SUBTITLE) that derive their font from the theme rather than storing
 * an explicit override in the named style definition.
 *
 * @param {Object}   doc      Full document object from Docs.Documents.get().
 * @param {Object[]} fontMap  Array of { oldFont, newFont } entries (FONT_MAP).
 * @returns {Object[]}        Array of updateNamedStyle request objects.
 */
function buildDocNamedStyleFontRequests(doc, fontMap) {
  const requests = [];
  const styles   = (doc.namedStyles && doc.namedStyles.styles) || [];

  // Determine NORMAL_TEXT font as fallback for styles with no explicit font.
  var normalTextWff    = null;
  var normalTextFamily = null;
  styles.forEach(function(ns) {
    if (ns.namedStyleType === "NORMAL_TEXT") {
      normalTextWff    = ns.textStyle && ns.textStyle.weightedFontFamily;
      normalTextFamily = normalTextWff
        ? normalTextWff.fontFamily
        : (ns.textStyle && ns.textStyle.fontFamily);
    }
  });

  styles.forEach(function(ns) {
    const wff            = ns.textStyle && ns.textStyle.weightedFontFamily;
    const explicitFamily = wff ? wff.fontFamily : (ns.textStyle && ns.textStyle.fontFamily);

    // For non-NORMAL_TEXT styles with no explicit font, proxy against
    // NORMAL_TEXT so theme-inherited heading fonts are still detected.
    const effectiveFamily = explicitFamily ||
      (ns.namedStyleType !== "NORMAL_TEXT" ? normalTextFamily : null);

    if (!effectiveFamily) return;

    var nsFontMatched = false;
    fontMap.forEach(function(mapping) {
      if (effectiveFamily === mapping.oldFont) {
        nsFontMatched = true;
        const weight = wff ? wff.weight
          : (ns.namedStyleType !== "NORMAL_TEXT" && normalTextWff ? normalTextWff.weight : 400);
        requests.push({
          updateNamedStyle: {
            namedStyle: {
              namedStyleType: ns.namedStyleType,
              textStyle: {
                weightedFontFamily: {
                  fontFamily: mapping.newFont,
                  weight:     weight,
                },
              },
            },
            fields: "textStyle.weightedFontFamily",
          },
        });
      }
    });

    // Replace any non-brand font not handled by FONT_MAP
    if (!nsFontMatched && BRAND_FONTS.indexOf(effectiveFamily) === -1) {
      const weight = wff ? wff.weight
        : (ns.namedStyleType !== "NORMAL_TEXT" && normalTextWff ? normalTextWff.weight : 400);
      requests.push({
        updateNamedStyle: {
          namedStyle: {
            namedStyleType: ns.namedStyleType,
            textStyle: {
              weightedFontFamily: {
                fontFamily: FALLBACK_FONT,
                weight:     weight,
              },
            },
          },
          fields: "textStyle.weightedFontFamily",
        },
      });
    }
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

  const nsLookup = buildNamedStyleLookup(doc);
  const requests = buildDocFontRequests(segments, FONT_MAP, nsLookup);
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

  function recordImageElements(el, segmentLabel) {
    if (!el.paragraph) return;
    (el.paragraph.elements || []).forEach(function(pe) {
      if (pe.inlineObjectElement) {
        objectSegment[pe.inlineObjectElement.inlineObjectId] = segmentLabel;
      }
    });
  }

  walkDocContent(doc.body ? doc.body.content : null, "body", recordImageElements);
  Object.keys(doc.headers || {}).forEach(function(id) {
    walkDocContent(doc.headers[id].content, "header:" + id, recordImageElements);
  });
  Object.keys(doc.footers || {}).forEach(function(id) {
    walkDocContent(doc.footers[id].content, "footer:" + id, recordImageElements);
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
// logDocTableStyles (diagnostic utility)
// ---------------------------------------------------------------------------

/**
 * Diagnostic utility — logs every table cell border and background color
 * found in the document, so you can see what values are actually stored
 * and whether getDocBorderRgb is returning the right values.
 *
 * Run this on a document whose borders weren't updated to understand what
 * color format is used (rgbColor vs themeColor vs absent).
 *
 * @param {string} docId
 */
function logDocTableStyles(docId) {
  const doc = Docs.Documents.get(docId);
  const BORDER_SIDES = ["borderLeft", "borderRight", "borderTop", "borderBottom"];

  function rgbToHex(rgb) {
    if (!rgb) return "(null)";
    var r = Math.round((rgb.red   || 0) * 255).toString(16).padStart(2, "0");
    var g = Math.round((rgb.green || 0) * 255).toString(16).padStart(2, "0");
    var b = Math.round((rgb.blue  || 0) * 255).toString(16).padStart(2, "0");
    return "#" + r + g + b;
  }

  var tableIndex = 0;

  function processTable(el) {
    if (!el.table) return;
    tableIndex++;
    Logger.log("=== Table %d (startIndex: %s) ===", tableIndex, el.startIndex);

    (el.table.tableRows || []).forEach(function(row, rowIdx) {
      (row.tableCells || []).forEach(function(cell, colIdx) {
        var ts = cell.tableCellStyle;
        if (!ts) {
          Logger.log("  [%d,%d] — no tableCellStyle", rowIdx, colIdx);
          return;
        }

        // Background
        var bgRgb = ts.backgroundColor && ts.backgroundColor.color && ts.backgroundColor.color.rgbColor;
        var bgTheme = ts.backgroundColor && ts.backgroundColor.color && ts.backgroundColor.color.themeColor;
        Logger.log(
          "  [%d,%d] bg: %s",
          rowIdx, colIdx,
          bgRgb ? rgbToHex(bgRgb) : (bgTheme ? "themeColor:" + bgTheme : "(none)")
        );

        // Borders
        BORDER_SIDES.forEach(function(side) {
          if (!ts[side]) return;
          var borderColor = ts[side].color && ts[side].color.color;
          var bRgb   = borderColor && borderColor.rgbColor;
          var bTheme = borderColor && borderColor.themeColor;
          var colorLabel = bRgb ? rgbToHex(bRgb) : (bTheme ? "themeColor:" + bTheme : "(none/transparent)");
          Logger.log(
            "  [%d,%d] %s: %s  width:%s  dash:%s",
            rowIdx, colIdx, side, colorLabel,
            (ts[side].width && ts[side].width.magnitude) || "?",
            ts[side].dashStyle || "?"
          );
        });
      });
    });
  }

  walkDocContent(doc.body ? doc.body.content : null, "", processTable);
  Logger.log("Done. Total tables found: %d", tableIndex);
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

  function collectImageElement(el, segmentId) {
    if (!el.paragraph) return;
    (el.paragraph.elements || []).forEach(function(pe) {
      if (pe.inlineObjectElement) {
        reverseIndex.push({
          objectId:   pe.inlineObjectElement.inlineObjectId,
          startIndex: pe.startIndex !== undefined ? pe.startIndex : 0,
          segmentId:  segmentId,
        });
      }
    });
  }

  walkDocContent(doc.body ? doc.body.content : null, "", collectImageElement);
  Object.keys(doc.headers || {}).forEach(function(id) {
    walkDocContent(doc.headers[id].content, id, collectImageElement);
  });
  Object.keys(doc.footers || {}).forEach(function(id) {
    walkDocContent(doc.footers[id].content, id, collectImageElement);
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
  const doc = Docs.Documents.get(docId);

  // Prefer an explicit direct URL from config; fall back to the Drive uc?id=
  // format which serves image bytes for publicly shared files without export.
  const newLogoUrl = LOGO_CONFIG.docsLogo.newLogoUrl ||
    ("https://drive.google.com/uc?id=" + LOGO_CONFIG.newLogoFileId);

  Logger.log("  replaceDocLogos: using image URL: %s", newLogoUrl);
  const requests = buildDocLogoRequests(doc, newLogoUrl, dryRun);

  if (dryRun) {
    Logger.log("  replaceDocLogos: dry run complete for %s", docId);
    return;
  }

  if (requests.length === 0) {
    Logger.log("  replaceDocLogos: no logo matches for %s", docId);
    return;
  }

  // Send via REST to avoid Advanced Service mis-serialising objectSize.
  batchUpdateDocWithUrlFetch(docId, requests);
  Logger.log("  replaceDocLogos: %d requests submitted for %s", requests.length, docId);
}

// ---------------------------------------------------------------------------
// Step 13 — updateDocsDocument (public orchestrator)
// ---------------------------------------------------------------------------

/**
 * Runs the full brand update pipeline on a single Google Doc:
 *   1. replaceDocColors  — explicit inline foreground color overrides
 *   2. replaceDocFonts   — Poppins / Figtree → Geist
 *   3. replaceDocLogos   — delete + re-insert logo images
 *
 * Each step can be called independently for isolated testing.
 * The dryRun flag is passed through to replaceDocLogos only.
 *
 * @param {string}  docId
 * @param {boolean} [dryRun=false]  Passed through to replaceDocLogos.
 */
function updateDocsDocument(docId, dryRun, options) {
  var opts = options || { colors: true, fonts: true, logo: true };
  Logger.log("Starting brand update for document: %s", docId);

  if (opts.colors) {
    try {
      replaceDocColors(docId);
    } catch (e) {
      Logger.log("  ERROR in replaceDocColors: %s", e.message);
      throw new Error("replaceDocColors failed: " + e.message);
    }
  }

  if (opts.fonts) {
    try {
      replaceDocFonts(docId);
    } catch (e) {
      Logger.log("  ERROR in replaceDocFonts: %s", e.message);
      throw new Error("replaceDocFonts failed: " + e.message);
    }
  }

  if (opts.logo) {
    try {
      replaceDocLogos(docId, dryRun);
    } catch (e) {
      Logger.log("  ERROR in replaceDocLogos: %s", e.message);
      throw new Error("replaceDocLogos failed: " + e.message);
    }
  }

  Logger.log("Done: %s", docId);
}

