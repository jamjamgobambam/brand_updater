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
  const matches       = findLogoMatches(doc, logoConfig, dryRun);

  if (dryRun) return [];
  if (matches.length === 0) return [];

  // Sort in reverse order so earlier inserts/deletes don't shift the
  // positions of later ones.
  matches.sort(function(a, b) { return b.startIndex - a.startIndex; });

  // Uniform scale factor for the replacement image (default 4x). Preserves
  // the new logo's aspect ratio inside the larger box.
  const scale = (logoConfig.scale && logoConfig.scale > 0) ? logoConfig.scale : 1;

  // Page-level fallback cap when the logo is NOT inside a table.
  const docStyle = doc.documentStyle || {};
  const pageW    = docStyle.pageSize    && docStyle.pageSize.width    && docStyle.pageSize.width.magnitude;
  const marginL  = docStyle.marginLeft  && docStyle.marginLeft.magnitude;
  const marginR  = docStyle.marginRight && docStyle.marginRight.magnitude;
  const pageMaxWidthPt = (pageW && marginL !== undefined && marginR !== undefined)
    ? (pageW - marginL - marginR)
    : null;

  // Global hard cap (brand max width). Applied alongside the cell/page
  // cap via Math.min, so a narrow cell still clamps tighter; this only
  // prevents the logo from getting too wide in roomy cells.
  const globalMaxPt = (typeof logoConfig.maxWidthInsertedIn === "number" && logoConfig.maxWidthInsertedIn > 0)
    ? logoConfig.maxWidthInsertedIn * 72
    : null;

  const requests = [];
  matches.forEach(function(match) {
    var newWidth  = match.widthPt  * scale;
    var newHeight = match.heightPt * scale;
    // Combine all applicable caps; cellWidthPt is preferred for in-table
    // logos, pageMaxWidthPt is the fallback. globalMaxPt always applies.
    var caps = [];
    if (typeof match.cellWidthPt === "number") caps.push(match.cellWidthPt);
    else if (pageMaxWidthPt !== null)          caps.push(pageMaxWidthPt);
    if (globalMaxPt !== null)                  caps.push(globalMaxPt);
    var maxWidthPt = caps.length ? Math.min.apply(null, caps) : null;
    if (maxWidthPt !== null && newWidth > maxWidthPt) {
      // Clamp width to the tightest cap; preserve aspect ratio.
      var ratio = maxWidthPt / newWidth;
      newWidth  = maxWidthPt;
      newHeight = newHeight * ratio;
    }

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
          width:  { magnitude: newWidth,  unit: "PT" },
          height: { magnitude: newHeight, unit: "PT" },
        },
      },
    });
  });

  return requests;
}

// ---------------------------------------------------------------------------
// findLogoMatches (internal)
// ---------------------------------------------------------------------------

/**
 * Walks a document and returns logo-image matches with their containing
 * table context (startIndex, segmentId, cellWidthPt, tableStartIndex,
 * colIndex, numColumns, firstColEmpty). Used by both buildDocLogoRequests
 * and buildDocTableStructureRequests.
 *
 * @param {Object}  doc         Full document from Docs.Documents.get().
 * @param {Object}  logoConfig  LOGO_CONFIG.docsLogo
 * @param {boolean} dryRun      If true, log each match and return [].
 * @returns {Object[]}          Array of match objects (empty when dryRun).
 */
function findLogoMatches(doc, logoConfig, dryRun) {
  const inlineObjects = doc.inlineObjects || {};
  const reverseIndex  = [];

  // Pre-compute, per table (keyed by tableStartIndex within a segmentId),
  // whether the first column is "empty" (no inline images and no
  // non-whitespace text in any of its row-0 cells). Detected during walk.
  const tableMeta = {};

  function isCellEmpty(cell) {
    var empty = true;
    (cell.content || []).forEach(function(el) {
      if (!empty) return;
      if (el.paragraph) {
        (el.paragraph.elements || []).forEach(function(pe) {
          if (!empty) return;
          if (pe.inlineObjectElement) { empty = false; return; }
          if (pe.textRun && pe.textRun.content && /\S/.test(pe.textRun.content)) {
            empty = false;
          }
        });
      }
      if (el.table) {
        // A nested table counts as content.
        empty = false;
      }
    });
    return empty;
  }

  function collectFromContent(contentArray, segmentId, tableCtx) {
    if (!contentArray) return;
    contentArray.forEach(function(el) {
      if (el.paragraph) {
        (el.paragraph.elements || []).forEach(function(pe) {
          if (pe.inlineObjectElement) {
            reverseIndex.push({
              objectId:        pe.inlineObjectElement.inlineObjectId,
              startIndex:      pe.startIndex !== undefined ? pe.startIndex : 0,
              segmentId:       segmentId,
              cellWidthPt:     tableCtx ? tableCtx.cellWidthPt : null,
              tableStartIndex: tableCtx ? tableCtx.tableStartIndex : null,
              rowIndex:        tableCtx ? tableCtx.rowIndex : null,
              colIndex:        tableCtx ? tableCtx.colIndex : null,
              numColumns:      tableCtx ? tableCtx.numColumns : null,
            });
          }
        });
      }
      if (el.table) {
        const tableStartIndex = el.startIndex !== undefined ? el.startIndex : 0;
        const colProps =
          (el.table.tableStyle && el.table.tableStyle.tableColumnProperties) || [];
        const numColumns = colProps.length || (el.table.columns || 0);

        // Determine if the first column is empty across all rows.
        var firstColEmpty = true;
        (el.table.tableRows || []).forEach(function(row) {
          if (!firstColEmpty) return;
          var cell0 = (row.tableCells || [])[0];
          if (!cell0 || !isCellEmpty(cell0)) firstColEmpty = false;
        });

        const metaKey = segmentId + "|" + tableStartIndex;
        tableMeta[metaKey] = {
          numColumns:    numColumns,
          firstColEmpty: firstColEmpty,
        };

        (el.table.tableRows || []).forEach(function(row, rowIndex) {
          (row.tableCells || []).forEach(function(cell, colIndex) {
            const colW =
              colProps[colIndex] &&
              colProps[colIndex].width &&
              colProps[colIndex].width.magnitude;
            const cs   = cell.tableCellStyle || {};
            const padL = (cs.paddingLeft  && cs.paddingLeft.magnitude)  || 0;
            const padR = (cs.paddingRight && cs.paddingRight.magnitude) || 0;
            const innerW = (typeof colW === "number") ? Math.max(0, colW - padL - padR) : null;
            collectFromContent(cell.content, segmentId, {
              cellWidthPt:     innerW,
              tableStartIndex: tableStartIndex,
              rowIndex:        rowIndex,
              colIndex:        colIndex,
              numColumns:      numColumns,
            });
          });
        });
      }
    });
  }

  collectFromContent(doc.body ? doc.body.content : null, "", null);
  Object.keys(doc.headers || {}).forEach(function(id) {
    collectFromContent(doc.headers[id].content, id, null);
  });
  Object.keys(doc.footers || {}).forEach(function(id) {
    collectFromContent(doc.footers[id].content, id, null);
  });

  const matches = [];
  // Build the set of known legacy logo sourceUris (legacy oldSourceUri +
  // current oldSourceUris[]). Membership in this set is the cheap fast path.
  const knownUris = {};
  if (logoConfig.oldSourceUri) knownUris[logoConfig.oldSourceUri] = true;
  (logoConfig.oldSourceUris || []).forEach(function(u) { if (u) knownUris[u] = true; });

  // Per-run cache of Gemini verdicts keyed by objectId so the same image
  // appearing multiple times only triggers one classifier call.
  const verdictByObjectId = {};

  reverseIndex.forEach(function(entry) {
    const inlineObj = inlineObjects[entry.objectId];
    if (!inlineObj) return;

    const embedded =
      inlineObj.inlineObjectProperties &&
      inlineObj.inlineObjectProperties.embeddedObject;
    if (!embedded) return;

    const imgProps  = embedded.imageProperties || {};
    const sourceUri = imgProps.sourceUri;
    const contentUri = imgProps.contentUri;
    const widthPt   = embedded.size && embedded.size.width  && embedded.size.width.magnitude;
    const heightPt  = embedded.size && embedded.size.height && embedded.size.height.magnitude;

    // Tier 1 — exact sourceUri match (deterministic, no API cost).
    let isMatch = !!(sourceUri && knownUris[sourceUri]);
    let matchReason = isMatch ? "sourceUri" : null;

    // Tier 2 — Gemini classifier on the image bytes. Only images
    // unambiguously classified as the brand logo become matches; this
    // prevents the size-based false positives we used to get on avatars
    // and unrelated content images.
    if (!isMatch && logoConfig.useGeminiClassifier !== false && contentUri) {
      var verdict = verdictByObjectId[entry.objectId];
      if (verdict === undefined) {
        try {
          const blob = fetchInlineObjectBlob_(contentUri);
          if (blob) {
            verdict = classifyLogo_(blob);
          } else {
            verdict = { is_logo: false, confidence: 0 };
          }
        } catch (e) {
          Logger.log("findLogoMatches: classifier failed for objectId=%s: %s",
                     entry.objectId, e && e.message);
          verdict = { is_logo: false, confidence: 0 };
        }
        verdictByObjectId[entry.objectId] = verdict;
      }
      if (typeof logoAction_ === "function" && logoAction_(verdict) === "replace") {
        isMatch = true;
        matchReason = "gemini(" + verdict.confidence.toFixed(2) + ")";
      }
    }

    if (!isMatch) return;

    if (dryRun) {
      Logger.log(
        "DRY RUN — logo match: objectId=%s segmentId=%s startIndex=%s reason=%s sourceUri=%s width=%sPT height=%sPT",
        entry.objectId, entry.segmentId, entry.startIndex,
        matchReason, sourceUri || "(null)", widthPt, heightPt
      );
      return;
    }

    const meta = (entry.tableStartIndex !== null && entry.tableStartIndex !== undefined)
      ? tableMeta[entry.segmentId + "|" + entry.tableStartIndex]
      : null;

    matches.push({
      startIndex:      entry.startIndex,
      segmentId:       entry.segmentId,
      widthPt:         widthPt,
      heightPt:        heightPt,
      cellWidthPt:     entry.cellWidthPt,
      tableStartIndex: entry.tableStartIndex,
      rowIndex:        entry.rowIndex,
      colIndex:        entry.colIndex,
      numColumns:      entry.numColumns,
      firstColEmpty:   meta ? meta.firstColEmpty : false,
    });
  });

  return matches;
}

// ---------------------------------------------------------------------------
// fetchInlineObjectBlob_ (internal)
// ---------------------------------------------------------------------------

/**
 * Fetch the bytes of an inline image's contentUri using the Apps Script
 * OAuth token. The contentUri is short-lived (~30 minutes) and requires
 * authentication; UrlFetchApp with a Bearer token is the supported way to
 * retrieve the image bytes from the Docs API.
 *
 * @param {string} contentUri
 * @returns {GoogleAppsScript.Base.Blob|null}
 */
function fetchInlineObjectBlob_(contentUri) {
  if (!contentUri) return null;
  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(contentUri, {
    method:             "get",
    headers:            { Authorization: "Bearer " + token },
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    Logger.log("fetchInlineObjectBlob_: HTTP %s for %s", code, contentUri);
    return null;
  }
  return resp.getBlob();
}

// ---------------------------------------------------------------------------
// logDocLogoCandidates (diagnostic, read-only)
// ---------------------------------------------------------------------------

/**
 * Lists every inline image in a document with its sourceUri, contentUri,
 * size, and (when useGemini=true) Gemini classifier verdict. Read-only —
 * use to discover which sourceUris belong to legitimate logos so you can
 * populate LOGO_CONFIG.docsLogo.oldSourceUris.
 *
 * @param {string}  docId
 * @param {boolean} [useGemini=false]  If true, run classifyLogo_ on each image.
 */
function logDocLogoCandidates(docId, useGemini) {
  const doc = Docs.Documents.get(docId);
  const inlineObjects = doc.inlineObjects || {};
  const ids = Object.keys(inlineObjects);
  Logger.log("logDocLogoCandidates: %d inline images in %s", ids.length, docId);

  ids.forEach(function(objectId) {
    const obj = inlineObjects[objectId];
    const embedded = obj.inlineObjectProperties && obj.inlineObjectProperties.embeddedObject;
    if (!embedded) return;
    const ip = embedded.imageProperties || {};
    const w  = embedded.size && embedded.size.width  && embedded.size.width.magnitude;
    const h  = embedded.size && embedded.size.height && embedded.size.height.magnitude;

    var verdictStr = "(skipped)";
    if (useGemini && ip.contentUri) {
      try {
        const blob = fetchInlineObjectBlob_(ip.contentUri);
        if (blob) {
          const v = classifyLogo_(blob);
          verdictStr = "is_logo=" + v.is_logo + " conf=" + (v.confidence || 0).toFixed(2) +
                       " action=" + (typeof logoAction_ === "function" ? logoAction_(v) : "?");
        } else {
          verdictStr = "(no blob)";
        }
      } catch (e) {
        verdictStr = "(error: " + (e && e.message) + ")";
      }
    }

    Logger.log(
      "  objectId=%s sizePT=%sx%s sourceUri=%s gemini=%s",
      objectId,
      w != null ? w.toFixed(1) : "?",
      h != null ? h.toFixed(1) : "?",
      ip.sourceUri || "(null)",
      verdictStr
    );
  });
}

// ---------------------------------------------------------------------------
// logDocLogoTables (diagnostic, read-only)
// ---------------------------------------------------------------------------

/**
 * Logs structural details for every table in the document that contains a
 * matched logo. Read-only — makes no API changes. Intended for diagnosing
 * layout issues before running the structural pass.
 *
 * For each affected table, logs:
 *   - segmentId, tableStartIndex, numColumns
 *   - page content width vs. sum of column widths
 *   - per-column: widthType, width (PT and inches)
 *   - per cell (row × col): empty?, has-image?, first ~50 chars of text
 *   - the matched logo's column index
 *
 * @param {string} docId
 */
function logDocLogoTables(docId) {
  const doc        = Docs.Documents.get(docId);
  const logoConfig = LOGO_CONFIG.docsLogo;
  const matches    = findLogoMatches(doc, logoConfig, false);

  if (matches.length === 0) {
    Logger.log("logDocLogoTables: no logo matches in %s", docId);
    return;
  }  // Page content width.
  const ds = doc.documentStyle || {};
  const pageW   = ds.pageSize    && ds.pageSize.width    && ds.pageSize.width.magnitude;
  const marginL = ds.marginLeft  && ds.marginLeft.magnitude;
  const marginR = ds.marginRight && ds.marginRight.magnitude;
  const pageContentPt = (pageW != null && marginL != null && marginR != null)
    ? (pageW - marginL - marginR) : null;

  // Group matches by (segmentId, tableStartIndex) so we don't double-log.
  const seen = {};
  matches.forEach(function(m) {
    if (m.tableStartIndex == null) return;
    const key = m.segmentId + "|" + m.tableStartIndex;
    if (seen[key]) return;
    seen[key] = true;

    // Find the actual table object so we can iterate rows/cells.
    const tableObj = findTableByStartIndex_(doc, m.segmentId, m.tableStartIndex);
    if (!tableObj) {
      Logger.log("logDocLogoTables: could not locate table at segmentId=%s startIndex=%s",
                 m.segmentId || "(body)", m.tableStartIndex);
      return;
    }

    const colProps = (tableObj.tableStyle && tableObj.tableStyle.tableColumnProperties) || [];
    var totalPt = 0;
    colProps.forEach(function(cp) {
      if (cp && cp.width && typeof cp.width.magnitude === "number") totalPt += cp.width.magnitude;
    });

    Logger.log("=== Table @ segmentId=%s tableStartIndex=%s ===",
               m.segmentId || "(body)", m.tableStartIndex);
    Logger.log("  numColumns=%s firstColEmpty=%s logoColIndex=%s",
               m.numColumns, m.firstColEmpty, m.colIndex);
    Logger.log("  pageContentPt=%s sumColPt=%s",
               pageContentPt, totalPt);
    colProps.forEach(function(cp, i) {
      const wt = cp && cp.widthType;
      const wp = cp && cp.width && cp.width.magnitude;
      Logger.log("  col[%d]: widthType=%s width=%sPT (%sin)",
                 i, wt || "(none)", wp != null ? wp.toFixed(2) : "(none)",
                 wp != null ? (wp / 72).toFixed(3) : "(none)");
    });

    (tableObj.tableRows || []).forEach(function(row, rIdx) {
      (row.tableCells || []).forEach(function(cell, cIdx) {
        var hasImage = false;
        var textPreview = "";
        (cell.content || []).forEach(function(el) {
          if (el.paragraph) {
            (el.paragraph.elements || []).forEach(function(pe) {
              if (pe.inlineObjectElement) hasImage = true;
              if (pe.textRun && pe.textRun.content) {
                textPreview += pe.textRun.content;
              }
            });
          }
        });
        const trimmed = textPreview.replace(/\s+/g, " ").trim();
        const empty = !hasImage && !/\S/.test(textPreview);
        Logger.log("  cell[r=%d,c=%d]: empty=%s hasImage=%s text=%s",
                   rIdx, cIdx, empty, hasImage,
                   JSON.stringify(trimmed.length > 50 ? trimmed.slice(0, 50) + "…" : trimmed));
      });
    });
  });
}

/**
 * Locate a table by its startIndex within a given segment.
 * @param {Object} doc
 * @param {string} segmentId  "" for body, header/footer ID otherwise.
 * @param {number} tableStartIndex
 * @returns {Object|null}     The table structural element's `.table`, or null.
 */
function findTableByStartIndex_(doc, segmentId, tableStartIndex) {
  var found = null;
  function scan(contentArray) {
    if (!contentArray || found) return;
    contentArray.forEach(function(el) {
      if (found) return;
      const startIndex = el.startIndex !== undefined ? el.startIndex : 0;
      if (el.table && startIndex === tableStartIndex) {
        found = el.table;
        return;
      }
      if (el.table) {
        (el.table.tableRows || []).forEach(function(row) {
          (row.tableCells || []).forEach(function(cell) {
            scan(cell.content);
          });
        });
      }
    });
  }

  if (segmentId === "" || segmentId == null) {
    scan(doc.body && doc.body.content);
  } else if (doc.headers && doc.headers[segmentId]) {
    scan(doc.headers[segmentId].content);
  } else if (doc.footers && doc.footers[segmentId]) {
    scan(doc.footers[segmentId].content);
  }
  return found;
}

// ---------------------------------------------------------------------------
// buildDocTableStructureRequests
// ---------------------------------------------------------------------------

/**
 * For every table that contains a matched logo, returns the structural
 * requests needed to bring the table into the new layout:
 *
 *   - If the table has 3 columns and the first column is empty across all
 *     rows (legacy spacer column), emit a deleteTableColumn(0) request.
 *   - Then emit one updateTableColumnProperties request per configured
 *     width in LOGO_CONFIG.docsLogo.tableColumnWidthsIn, applied to the
 *     POST-deletion column indices.
 *
 * These requests must be sent in a separate batchUpdate before logo edits,
 * because deleteTableColumn shifts the document's text indexes and would
 * invalidate the cached logo startIndex values.
 *
 * @param {Object} doc
 * @returns {Object[]}  Array of request objects (may be empty).
 */
function buildDocTableStructureRequests(doc) {
  const logoConfig = LOGO_CONFIG.docsLogo;
  const widthsIn   = logoConfig.tableColumnWidthsIn || [];
  const indentIn   = logoConfig.firstColumnTextIndentIn;
  const indentPt   = (typeof indentIn === "number" && indentIn > 0) ? indentIn * 72 : null;
  const matches    = findLogoMatches(doc, logoConfig, false);

  if (matches.length === 0) return [];

  // De-duplicate by table; build per-table request groups so they can be
  // emitted in reverse-tableStartIndex order. Multiple legacy logo tables
  // in the same segment shift each other's indexes when the earlier one
  // has a column deleted; emitting later tables first sidesteps this.
  const seenTables = {};
  const groups     = [];

  function cellHasText(cell) {
    var hasText = false;
    (cell.content || []).forEach(function(el) {
      if (hasText || !el.paragraph) return;
      (el.paragraph.elements || []).forEach(function(pe) {
        if (hasText) return;
        if (pe.textRun && pe.textRun.content && /\S/.test(pe.textRun.content)) {
          hasText = true;
        }
      });
    });
    return hasText;
  }

  function paragraphIndentRequests(cell, segmentId, magnitudePt) {
    var out = [];
    (cell.content || []).forEach(function(el) {
      if (!el.paragraph) return;
      const startIndex = el.startIndex !== undefined ? el.startIndex : 0;
      const endIndex   = el.endIndex;
      if (endIndex === undefined) return;
      out.push({
        updateParagraphStyle: {
          range: {
            startIndex: startIndex,
            endIndex:   endIndex,
            segmentId:  segmentId,
          },
          paragraphStyle: {
            indentStart:     { magnitude: magnitudePt, unit: "PT" },
            indentFirstLine: { magnitude: magnitudePt, unit: "PT" },
          },
          fields: "indentStart,indentFirstLine",
        },
      });
    });
    return out;
  }

  matches.forEach(function(match) {
    if (match.tableStartIndex === null || match.tableStartIndex === undefined) return;
    const key = match.segmentId + "|" + match.tableStartIndex;
    if (seenTables[key]) return;
    seenTables[key] = true;

    // Strict legacy-table discriminator. Only restructure tables that
    // visibly look like the legacy [empty | title | logo] footer/header:
    //   - exactly 3 columns
    //   - col 0 empty across all rows
    //   - matched logo sits in the LAST column
    //   - col 1 of the matched logo's row contains non-whitespace text
    //     (i.e., title text — distinguishes from generic 3-col image
    //     tables that happen to have an empty leading column).
    if (match.numColumns !== 3) return;
    if (!match.firstColEmpty) return;
    if (match.colIndex !== match.numColumns - 1) return;

    const tableObj = findTableByStartIndex_(doc, match.segmentId, match.tableStartIndex);
    if (!tableObj) return;
    const matchedRow = (tableObj.tableRows || [])[match.rowIndex];
    const middleCell = matchedRow && (matchedRow.tableCells || [])[1];
    if (!middleCell || !cellHasText(middleCell)) return;

    const groupRequests = [];

    // Apply paragraph indents using PRE-deletion text indexes. After
    // deleteTableColumn (which removes col 0), what was col 1 becomes col 0
    // and what was col 2 becomes col 1. updateParagraphStyle does not shift
    // text indexes, so emitting these BEFORE the deletion in the same
    // batchUpdate is safe.
    if (indentPt !== null) {
      (tableObj.tableRows || []).forEach(function(row) {
        const cells = row.tableCells || [];
        if (cells[1]) {
          // Old col 1 → new col 0 (title): apply configured indent.
          paragraphIndentRequests(cells[1], match.segmentId, indentPt)
            .forEach(function(r) { groupRequests.push(r); });
        }
        if (cells[2]) {
          // Old col 2 → new col 1 (logo): zero indents so the logo image
          // sits flush at the cell's left edge.
          paragraphIndentRequests(cells[2], match.segmentId, 0)
            .forEach(function(r) { groupRequests.push(r); });
        }
      });
    }

    groupRequests.push({
      deleteTableColumn: {
        tableCellLocation: {
          tableStartLocation: { index: match.tableStartIndex, segmentId: match.segmentId },
          rowIndex:           0,
          columnIndex:        0,
        },
      },
    });

    if (widthsIn.length > 0) {
      const remainingCols = match.numColumns - 1;
      widthsIn.forEach(function(inches, colIdx) {
        if (colIdx >= remainingCols) return;
        groupRequests.push({
          updateTableColumnProperties: {
            tableStartLocation: { index: match.tableStartIndex, segmentId: match.segmentId },
            columnIndices:      [colIdx],
            tableColumnProperties: {
              widthType: "FIXED_WIDTH",
              width:     { magnitude: inches * 72, unit: "PT" },
            },
            fields: "widthType,width",
          },
        });
      });
    }

    groups.push({
      segmentId:       match.segmentId,
      tableStartIndex: match.tableStartIndex,
      requests:        groupRequests,
    });
  });

  // Sort groups by tableStartIndex DESC so structural edits to later
  // tables happen before any deleteTableColumn shifts earlier indexes.
  groups.sort(function(a, b) { return b.tableStartIndex - a.tableStartIndex; });

  const requests = [];
  groups.forEach(function(g) {
    g.requests.forEach(function(r) { requests.push(r); });
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
  var doc = Docs.Documents.get(docId);

  // Prefer an explicit direct URL from config; otherwise build a direct
  // content URL from the CODEAI_LOGO_DRIVE_ID Script Property.
  //
  // The Docs API cannot follow Drive redirects, so the legacy
  // `drive.google.com/uc?id=…` form returns 500 intermittently. The
  // `lh3.googleusercontent.com/d/{id}` host is the same CDN Drive itself
  // serves images from and returns the bytes directly — provided the file
  // is shared as "Anyone with the link can view".
  const newLogoUrl = LOGO_CONFIG.docsLogo.newLogoUrl ||
    ("https://lh3.googleusercontent.com/d/" + getCodeAILogoFileId_());

  Logger.log("  replaceDocLogos: using image URL: %s", newLogoUrl);

  if (dryRun) {
    buildDocLogoRequests(doc, newLogoUrl, true);
    Logger.log("  replaceDocLogos: dry run complete for %s", docId);
    return;
  }

  // Pass 1 — structural table changes (delete legacy spacer column, set
  // configured column widths). These shift document text indexes, so we
  // must send them in a separate batch BEFORE building logo edits, and
  // re-fetch the doc afterwards so logo startIndexes are accurate.
  // Gated by LOGO_CONFIG.docsLogo.restructureLegacyTable so this pass can
  // be disabled while diagnosing layout issues.
  if (LOGO_CONFIG.docsLogo.restructureLegacyTable !== false) {
    const structureRequests = buildDocTableStructureRequests(doc);
    if (structureRequests.length > 0) {
      batchUpdateDocWithUrlFetch(docId, structureRequests);
      Logger.log("  replaceDocLogos: %d table-structure requests submitted for %s",
                 structureRequests.length, docId);
      doc = Docs.Documents.get(docId);
    }
  } else {
    Logger.log("  replaceDocLogos: restructureLegacyTable=false — skipping table-structure pass");
  }

  // Pass 2 — logo delete + insert against fresh indexes.
  const logoRequests = buildDocLogoRequests(doc, newLogoUrl, false);
  if (logoRequests.length === 0) {
    Logger.log("  replaceDocLogos: no logo matches for %s", docId);
    return;
  }

  batchUpdateDocWithUrlFetch(docId, logoRequests);
  Logger.log("  replaceDocLogos: %d logo requests submitted for %s", logoRequests.length, docId);
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

