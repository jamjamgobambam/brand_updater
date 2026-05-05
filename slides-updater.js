// =============================================================================
// slides-updater.js — Google Slides brand updater (colors, fonts, logos)
// Depends on globals defined in utils.js: COLOR_MAP, FONT_MAP, LOGO_CONFIG,
// hexToNormalizedRgb, normalizedRgbMatches, HYPERLINK_NEW_HEX
// =============================================================================

// ---------------------------------------------------------------------------
// Step 7 — updateMasterThemeColors
// ---------------------------------------------------------------------------

/**
 * Updates the ColorScheme on every master slide by replacing only the
 * Accent 1–6, HYPERLINK, and FOLLOWED_HYPERLINK slots per COLOR_MAP.
 * DARK1, DARK2, LIGHT1, LIGHT2 are carried through unchanged.
 *
 * @param {string} presentationId
 * @param {Object[]} masters  Array of master page objects from the Slides API.
 */
function updateMasterThemeColors(presentationId, masters) {
  const requests = [];

  // New hex per Accent slot, sourced from the dedicated ACCENT_NEW_HEXES constant
  // rather than COLOR_MAP, which now has multiple entries per color family.
  const accentTypes = ["ACCENT1", "ACCENT2", "ACCENT3", "ACCENT4", "ACCENT5", "ACCENT6"];
  const accentNewHexes = ACCENT_NEW_HEXES;

  masters.forEach(function(master) {
    const existingColors =
      master.pageProperties &&
      master.pageProperties.colorScheme &&
      master.pageProperties.colorScheme.colors;

    if (!existingColors) return;

    // Deep-copy then patch only the target slots.
    // NOTE: ThemeColorPair.color is a bare RgbColor ({red, green, blue}) —
    // not an OpaqueColor wrapper. Do NOT wrap in { rgbColor: ... } here.
    const updatedColors = existingColors.map(function(entry) {
      const type = entry.type;
      const accentIndex = accentTypes.indexOf(type);
      if (accentIndex !== -1) {
        return {
          type: type,
          color: hexToNormalizedRgb(accentNewHexes[accentIndex]),
        };
      }
      if (type === "HYPERLINK" || type === "FOLLOWED_HYPERLINK") {
        return {
          type: type,
          color: hexToNormalizedRgb(HYPERLINK_NEW_HEX),
        };
      }
      // DARK1, DARK2, LIGHT1, LIGHT2 — preserve unchanged
      return entry;
    });

    requests.push({
      updatePageProperties: {
        objectId: master.objectId,
        pageProperties: {
          colorScheme: { colors: updatedColors },
        },
        fields: "colorScheme",
      },
    });
  });

  if (requests.length === 0) return;
  Slides.Presentations.batchUpdate({ requests: requests }, presentationId);
}

// ---------------------------------------------------------------------------
// Step 8 — buildInlineColorRequests
// ---------------------------------------------------------------------------

/**
 * Traverses all pages and builds batchUpdate request objects for every
 * inline (direct) RGB color that matches an entry in colorMap.
 *
 * Covers: page background, shape fill, shape outline, text run foreground,
 * table cell background fill, and line fill.
 *
 * @param {Object[]} pages     All page objects (masters + layouts + slides).
 * @param {Object[]} colorMap  Array of { oldHex, newHex } pairs.
 * @returns {Object[]}         Array of batchUpdate request objects.
 */
function buildInlineColorRequests(pages, colorMap) {
  const requests = [];

  pages.forEach(function(page) {
    const pageId = page.objectId;

    // --- Page background ---
    const bgRgb =
      page.pageProperties &&
      page.pageProperties.pageBackgroundFill &&
      page.pageProperties.pageBackgroundFill.solidFill &&
      page.pageProperties.pageBackgroundFill.solidFill.color &&
      page.pageProperties.pageBackgroundFill.solidFill.color.rgbColor;

    if (bgRgb) {
      var bgNewHex = findColorMapping(bgRgb, colorMap, COLOR_DISTANCE_THRESHOLD);
      if (bgNewHex) {
        requests.push({
          updatePageProperties: {
            objectId: pageId,
            pageProperties: {
              pageBackgroundFill: {
                solidFill: {
                  color: { rgbColor: hexToNormalizedRgb(bgNewHex) },
                },
              },
            },
            fields: "pageBackgroundFill.solidFill.color",
          },
        });
      }
    }

    // --- Page elements ---
    (page.pageElements || []).forEach(function(element) {
      const eid = element.objectId;

      // Shape fill
      const shapeFillRgb =
        element.shape &&
        element.shape.shapeProperties &&
        element.shape.shapeProperties.shapeBackgroundFill &&
        element.shape.shapeProperties.shapeBackgroundFill.solidFill &&
        element.shape.shapeProperties.shapeBackgroundFill.solidFill.color &&
        element.shape.shapeProperties.shapeBackgroundFill.solidFill.color.rgbColor;

      if (shapeFillRgb) {
        var shapeFillNewHex = findColorMapping(shapeFillRgb, colorMap, COLOR_DISTANCE_THRESHOLD);
        if (shapeFillNewHex) {
          requests.push({
            updateShapeProperties: {
              objectId: eid,
              shapeProperties: {
                shapeBackgroundFill: {
                  solidFill: {
                    color: { rgbColor: hexToNormalizedRgb(shapeFillNewHex) },
                  },
                },
              },
              fields: "shapeBackgroundFill.solidFill.color",
            },
          });
        }
      }

      // Shape outline
      const outlineRgb =
        element.shape &&
        element.shape.shapeProperties &&
        element.shape.shapeProperties.outline &&
        element.shape.shapeProperties.outline.outlineFill &&
        element.shape.shapeProperties.outline.outlineFill.solidFill &&
        element.shape.shapeProperties.outline.outlineFill.solidFill.color &&
        element.shape.shapeProperties.outline.outlineFill.solidFill.color.rgbColor;

      if (outlineRgb) {
        var outlineNewHex = findColorMapping(outlineRgb, colorMap, COLOR_DISTANCE_THRESHOLD);
        if (outlineNewHex) {
          requests.push({
            updateShapeProperties: {
              objectId: eid,
              shapeProperties: {
                outline: {
                  outlineFill: {
                    solidFill: {
                      color: { rgbColor: hexToNormalizedRgb(outlineNewHex) },
                    },
                  },
                },
              },
              fields: "outline.outlineFill.solidFill.color",
            },
          });
        }
      }

      // Text run foreground colors
      const textElements =
        element.shape &&
        element.shape.text &&
        element.shape.text.textElements;

      if (textElements) {
        textElements.forEach(function(te) {
          if (!te.textRun) return;
          const fgRgb =
            te.textRun.style &&
            te.textRun.style.foregroundColor &&
            te.textRun.style.foregroundColor.opaqueColor &&
            te.textRun.style.foregroundColor.opaqueColor.rgbColor;

          if (!fgRgb) return;
          var fgNewHex = findColorMapping(fgRgb, colorMap, COLOR_DISTANCE_THRESHOLD);
          if (fgNewHex) {
            requests.push({
              updateTextStyle: {
                objectId: eid,
                textRange: {
                  type: "FIXED_RANGE",
                  startIndex: te.startIndex !== undefined ? te.startIndex : 0,
                  endIndex: te.endIndex,
                },
                style: {
                  foregroundColor: {
                    opaqueColor: {
                      rgbColor: hexToNormalizedRgb(fgNewHex),
                    },
                  },
                },
                fields: "foregroundColor",
              },
            });
          }
        });
      }

      // Table cell background fill
      const tableRows =
        element.table &&
        element.table.tableRows;

      if (tableRows) {
        tableRows.forEach(function(row) {
          (row.tableCells || []).forEach(function(cell) {
            const cellRgb =
              cell.tableCellProperties &&
              cell.tableCellProperties.tableCellBackgroundFill &&
              cell.tableCellProperties.tableCellBackgroundFill.solidFill &&
              cell.tableCellProperties.tableCellBackgroundFill.solidFill.color &&
              cell.tableCellProperties.tableCellBackgroundFill.solidFill.color.rgbColor;

            if (cellRgb) {
              var cellNewHex = findColorMapping(cellRgb, colorMap, COLOR_DISTANCE_THRESHOLD);
              if (cellNewHex) {
                requests.push({
                  updateTableCellProperties: {
                    objectId: eid,
                    tableRange: {
                      location: {
                        rowIndex: cell.location.rowIndex,
                        columnIndex: cell.location.columnIndex,
                      },
                      rowSpan: 1,
                      columnSpan: 1,
                    },
                    tableCellProperties: {
                      tableCellBackgroundFill: {
                        solidFill: {
                          color: { rgbColor: hexToNormalizedRgb(cellNewHex) },
                        },
                      },
                    },
                    fields: "tableCellBackgroundFill.solidFill.color",
                  },
                });
              }
            }

            // Table cell text foreground colors
            const cellTextElements = cell.text && cell.text.textElements;
            if (cellTextElements) {
              cellTextElements.forEach(function(te) {
                if (!te.textRun) return;
                const fgRgb =
                  te.textRun.style &&
                  te.textRun.style.foregroundColor &&
                  te.textRun.style.foregroundColor.opaqueColor &&
                  te.textRun.style.foregroundColor.opaqueColor.rgbColor;
                if (!fgRgb) return;
                var fgNewHex = findColorMapping(fgRgb, colorMap, COLOR_DISTANCE_THRESHOLD);
                if (fgNewHex) {
                  requests.push({
                    updateTextStyle: {
                      objectId: eid,
                      cellLocation: {
                        rowIndex: cell.location.rowIndex,
                        columnIndex: cell.location.columnIndex,
                      },
                      textRange: {
                        type: "FIXED_RANGE",
                        startIndex: te.startIndex !== undefined ? te.startIndex : 0,
                        endIndex: te.endIndex,
                      },
                      style: {
                        foregroundColor: {
                          opaqueColor: { rgbColor: hexToNormalizedRgb(fgNewHex) },
                        },
                      },
                      fields: "foregroundColor",
                    },
                  });
                }
              });
            }
          });
        });
      }

      // Line fill
      const lineRgb =
        element.line &&
        element.line.lineProperties &&
        element.line.lineProperties.lineFill &&
        element.line.lineProperties.lineFill.solidFill &&
        element.line.lineProperties.lineFill.solidFill.color &&
        element.line.lineProperties.lineFill.solidFill.color.rgbColor;

      if (lineRgb) {
        var lineNewHex = findColorMapping(lineRgb, colorMap, COLOR_DISTANCE_THRESHOLD);
        if (lineNewHex) {
          requests.push({
            updateLineProperties: {
              objectId: eid,
              lineProperties: {
                lineFill: {
                  solidFill: {
                    color: { rgbColor: hexToNormalizedRgb(lineNewHex) },
                  },
                },
              },
              fields: "lineFill.solidFill.color",
            },
          });
        }
      }
    });
  });

  return requests;
}

// ---------------------------------------------------------------------------
// Step 9 — replaceInlineColors
// ---------------------------------------------------------------------------

/**
 * Replaces all inline RGB colors matching COLOR_MAP across masters, layouts,
 * and slides. Splits requests into batches to stay under the 500-request API limit.
 *
 * @param {string} presentationId
 * @param {Object} [cachedPresentation]  Pre-fetched presentation object; fetched if omitted.
 */
function replaceInlineColors(presentationId, cachedPresentation) {
  const presentation = cachedPresentation || getPresentation(presentationId);
  const allPages = [].concat(
    presentation.masters  || [],
    presentation.layouts  || [],
    presentation.slides   || []
  );

  const requests = buildInlineColorRequests(allPages, COLOR_MAP);
  if (requests.length === 0) return;

  // Batch in chunks of 480 to avoid hitting the 500-request API limit
  const BATCH_SIZE = 480;
  for (let i = 0; i < requests.length; i += BATCH_SIZE) {
    const chunk = requests.slice(i, i + BATCH_SIZE);
    Slides.Presentations.batchUpdate({ requests: chunk }, presentationId);
  }
}

// ---------------------------------------------------------------------------
// Step 12 — buildFontRequests
// ---------------------------------------------------------------------------

/**
 * Traverses all pages and builds updateTextStyle request objects for every
 * text run whose explicit font matches an entry in fontMap.
 * Preserves font weight via weightedFontFamily. Runs with a null fontFamily
 * (inheriting from the master) are intentionally skipped.
 *
 * @param {Object[]} pages    All page objects (masters + layouts + slides).
 * @param {Object[]} fontMap  Array of { oldFont, newFont } pairs.
 * @returns {Object[]}        Array of updateTextStyle request objects.
 */
function buildFontRequests(pages, fontMap) {
  const requests = [];

  pages.forEach(function(page) {
    (page.pageElements || []).forEach(function(element) {
      const eid = element.objectId;

      // Helper: builds font requests for a single text elements array,
      // optionally scoped to a specific table cell via cellLocation.
      function processFontTextElements(textElements, cellLocation) {
        if (!textElements) return;
        textElements.forEach(function(te) {
          if (!te.textRun) return;
          const style = te.textRun.style || {};

          // Prefer weightedFontFamily (includes weight), fall back to fontFamily
          const wff = style.weightedFontFamily;
          const fontFamily = wff ? wff.fontFamily : style.fontFamily;
          if (!fontFamily) return; // null means inheriting — leave untouched

          var fontMatched = false;
          fontMap.forEach(function(mapping) {
            if (fontFamily === mapping.oldFont) {
              fontMatched = true;
              const existingWeight = wff ? wff.weight : 400;
              var req = {
                objectId: eid,
                textRange: {
                  type: "FIXED_RANGE",
                  startIndex: te.startIndex !== undefined ? te.startIndex : 0,
                  endIndex: te.endIndex,
                },
                style: {
                  weightedFontFamily: {
                    fontFamily: mapping.newFont,
                    weight: existingWeight,
                  },
                },
                fields: "weightedFontFamily",
              };
              if (cellLocation) req.cellLocation = cellLocation;
              requests.push({ updateTextStyle: req });
            }
          });

          // Replace any non-brand font not already handled by FONT_MAP
          if (!fontMatched && BRAND_FONTS.indexOf(fontFamily) === -1) {
            const existingWeight = wff ? wff.weight : 400;
            var req = {
              objectId: eid,
              textRange: {
                type: "FIXED_RANGE",
                startIndex: te.startIndex !== undefined ? te.startIndex : 0,
                endIndex: te.endIndex,
              },
              style: {
                weightedFontFamily: {
                  fontFamily: FALLBACK_FONT,
                  weight: existingWeight,
                },
              },
              fields: "weightedFontFamily",
            };
            if (cellLocation) req.cellLocation = cellLocation;
            requests.push({ updateTextStyle: req });
          }
        });
      }

      // Shape text
      processFontTextElements(
        element.shape && element.shape.text && element.shape.text.textElements,
        null
      );

      // Table cell text
      const tableRows = element.table && element.table.tableRows;
      if (tableRows) {
        tableRows.forEach(function(row) {
          (row.tableCells || []).forEach(function(cell) {
            processFontTextElements(
              cell.text && cell.text.textElements,
              { rowIndex: cell.location.rowIndex, columnIndex: cell.location.columnIndex }
            );
          });
        });
      }
    });
  });

  return requests;
}

// ---------------------------------------------------------------------------
// Step 13 — replaceFonts
// ---------------------------------------------------------------------------

/**
 * Replaces all explicit Poppins/Figtree font references with Lexend across
 * masters, layouts, and slides. Splits requests into batches.
 *
 * @param {string} presentationId
 * @param {Object} [cachedPresentation]  Pre-fetched presentation object; fetched if omitted.
 */
function replaceFonts(presentationId, cachedPresentation) {
  const presentation = cachedPresentation || getPresentation(presentationId);
  const allPages = [].concat(
    presentation.masters  || [],
    presentation.layouts  || [],
    presentation.slides   || []
  );

  // Include speaker notes pages from each slide
  (presentation.slides || []).forEach(function(slide) {
    if (slide.slideProperties && slide.slideProperties.notesPage) {
      allPages.push(slide.slideProperties.notesPage);
    }
  });

  const requests = buildFontRequests(allPages, FONT_MAP);
  if (requests.length === 0) return;

  const BATCH_SIZE = 480;
  for (let i = 0; i < requests.length; i += BATCH_SIZE) {
    const chunk = requests.slice(i, i + BATCH_SIZE);
    Slides.Presentations.batchUpdate({ requests: chunk }, presentationId);
  }
}

// ---------------------------------------------------------------------------
// logPresentationColors (diagnostic utility)
// ---------------------------------------------------------------------------

/**
 * Diagnostic utility — logs every inline RGB color found in the presentation
 * alongside its Euclidean distance to each old brand color in COLOR_MAP.
 * Run this on a presentation whose colors weren't changed to understand what
 * values are actually stored and whether the distance threshold is too low.
 *
 * Also logs the master theme color scheme so you can confirm whether shapes
 * use theme-referenced colors (handled by updateMasterThemeColors) or inline
 * RGB values (handled by buildInlineColorRequests).
 *
 * @param {string} presentationId
 */
function logPresentationColors(presentationId) {
  const presentation = getPresentation(presentationId);

  // --- Master theme color scheme ---
  Logger.log("=== MASTER THEME COLOR SCHEME ===");
  (presentation.masters || []).forEach(function(master) {
    const colors =
      master.pageProperties &&
      master.pageProperties.colorScheme &&
      master.pageProperties.colorScheme.colors;
    if (!colors) return;
    colors.forEach(function(entry) {
      const c = entry.color;
      const rgb = c && c.rgbColor;
      if (rgb) {
        const hex = "#" +
          Math.round((rgb.red   || 0) * 255).toString(16).padStart(2, "0") +
          Math.round((rgb.green || 0) * 255).toString(16).padStart(2, "0") +
          Math.round((rgb.blue  || 0) * 255).toString(16).padStart(2, "0");
        Logger.log("  %s → %s (rgbColor)", entry.type, hex);
      } else if (c && c.themeColor) {
        Logger.log("  %s → themeColor:%s", entry.type, c.themeColor);
      }
    });
  });

  // --- Inline RGB colors ---
  Logger.log("=== INLINE RGB COLORS ===");
  const allPages = [].concat(
    presentation.masters  || [],
    presentation.layouts  || [],
    presentation.slides   || []
  );

  var found = 0;
  allPages.forEach(function(page) {
    var pageLabel = (page.pageProperties && page.pageProperties.name) || page.objectId;

    function logRgb(rgb, location) {
      if (!rgb) return;
      found++;
      const hex = "#" +
        Math.round((rgb.red   || 0) * 255).toString(16).padStart(2, "0") +
        Math.round((rgb.green || 0) * 255).toString(16).padStart(2, "0") +
        Math.round((rgb.blue  || 0) * 255).toString(16).padStart(2, "0");
      var closest = null, closestDist = Infinity;
      COLOR_MAP.forEach(function(m) {
        var d = colorDistance(rgb, m.oldHex);
        if (d < closestDist) { closestDist = d; closest = m.oldHex; }
      });
      var withinThreshold = closestDist <= COLOR_DISTANCE_THRESHOLD;
      Logger.log(
        "  [%s] %s — hex:%s | closest old brand color:%s | dist:%.1f | within threshold:%s",
        pageLabel, location, hex, closest, closestDist, withinThreshold
      );
    }

    var bgRgb =
      page.pageProperties &&
      page.pageProperties.pageBackgroundFill &&
      page.pageProperties.pageBackgroundFill.solidFill &&
      page.pageProperties.pageBackgroundFill.solidFill.color &&
      page.pageProperties.pageBackgroundFill.solidFill.color.rgbColor;
    logRgb(bgRgb, "page background");

    (page.pageElements || []).forEach(function(el) {
      var fillRgb =
        el.shape && el.shape.shapeProperties &&
        el.shape.shapeProperties.shapeBackgroundFill &&
        el.shape.shapeProperties.shapeBackgroundFill.solidFill &&
        el.shape.shapeProperties.shapeBackgroundFill.solidFill.color &&
        el.shape.shapeProperties.shapeBackgroundFill.solidFill.color.rgbColor;
      logRgb(fillRgb, "shape fill " + el.objectId);

      var fillTheme =
        el.shape && el.shape.shapeProperties &&
        el.shape.shapeProperties.shapeBackgroundFill &&
        el.shape.shapeProperties.shapeBackgroundFill.solidFill &&
        el.shape.shapeProperties.shapeBackgroundFill.solidFill.color &&
        el.shape.shapeProperties.shapeBackgroundFill.solidFill.color.themeColor;
      if (fillTheme) {
        Logger.log("  [%s] shape fill %s — themeColor:%s (handled by master theme update)", pageLabel, el.objectId, fillTheme);
      }
    });
  });

  if (found === 0) {
    Logger.log("  (no inline rgbColor fills found — shapes likely use theme color references)");
  }
  Logger.log("=== DONE (found %d inline RGB colors) ===", found);
}

// ---------------------------------------------------------------------------
// Step 14 — logAllImages (diagnostic utility)
// ---------------------------------------------------------------------------

/**
 * Logs details of every image element on master, layout, AND slide pages.
 * Run once on a representative presentation to:
 *   1. Identify a stable substring of the existing logo's contentUrl/sourceUrl
 *      to copy into LOGO_CONFIG.slidesLogo.oldContentUrlSubstrings.
 *   2. Verify that each logo's center falls inside one of the configured zones.
 *
 * @param {string} presentationId
 */
function logAllImages(presentationId) {
  const presentation = Slides.Presentations.get(presentationId);
  const pageWidth  = presentation.pageSize.width.magnitude;
  const pageHeight = presentation.pageSize.height.magnitude;

  const pages = [].concat(
    (presentation.masters || []).map(function(p) { return { page: p, kind: "master" }; }),
    (presentation.layouts || []).map(function(p) { return { page: p, kind: "layout" }; }),
    (presentation.slides  || []).map(function(p) { return { page: p, kind: "slide"  }; })
  );

  pages.forEach(function(entry) {
    const page = entry.page;
    const pageName = page.pageProperties && page.pageProperties.name
      ? page.pageProperties.name
      : page.objectId;

    (page.pageElements || []).forEach(function(element) {
      if (!element.image) return;
      if (!element.transform) {
        Logger.log("[%s] Image [%s] on page [%s]: no transform (at default position)",
          entry.kind, element.objectId, pageName);
        return;
      }

      const tx = element.transform.translateX || 0;
      const ty = element.transform.translateY || 0;
      const w  = element.size.width.magnitude;
      const h  = element.size.height.magnitude;

      const centerX = (tx + w / 2) / pageWidth;
      const centerY = (ty + h / 2) / pageHeight;
      const widthPct  = w / pageWidth;
      const heightPct = h / pageHeight;

      // Determine which configured zone (if any) this image's center falls in.
      const zones = (LOGO_CONFIG.slidesLogo && LOGO_CONFIG.slidesLogo.zones) || [];
      var zoneHit = "(none)";
      for (var i = 0; i < zones.length; i++) {
        const z = zones[i];
        if (centerX >= z.xMin && centerX <= z.xMax &&
            centerY >= z.yMin && centerY <= z.yMax) {
          zoneHit = z.name;
          break;
        }
      }

      Logger.log(
        "[%s] Image [%s] on page [%s]: centerX=%.3f centerY=%.3f w=%.3f h=%.3f zone=%s\n  contentUrl=%s\n  sourceUrl=%s",
        entry.kind,
        element.objectId,
        pageName,
        centerX,
        centerY,
        widthPct,
        heightPct,
        zoneHit,
        element.image.contentUrl || "(none)",
        element.image.sourceUrl  || "(none)"
      );
    });
  });
}

// ---------------------------------------------------------------------------
// Step 17 — classifyLogoElement
// ---------------------------------------------------------------------------

/**
 * Layered logo detection. Returns null if not a logo, otherwise a match
 * descriptor: { matchedBy: "contentUrl"|"zone", zoneName?: string }.
 *
 * Primary match — image source URL substring:
 *   If LOGO_CONFIG.slidesLogo.oldContentUrlSubstrings is non-empty AND any
 *   substring appears in element.image.contentUrl or element.image.sourceUrl,
 *   the element is a logo regardless of position. This match bypasses the
 *   sizeBounds filter.
 *
 * Fallback match — position zone + size/aspect filter:
 *   If the element's center falls inside any zone in LOGO_CONFIG.slidesLogo.zones
 *   AND the element passes LOGO_CONFIG.slidesLogo.sizeBounds (width/height as
 *   fraction of slide dims, aspect = width/height), the element is a logo.
 *
 * @param {Object} element     A page element from the Slides API.
 * @param {number} pageWidth   Slide width in EMUs.
 * @param {number} pageHeight  Slide height in EMUs.
 * @returns {{matchedBy: string, zoneName?: string} | null}
 */
function classifyLogoElement(element, pageWidth, pageHeight) {
  if (!element.image)     return null;
  if (!element.transform) return null;

  const cfg = LOGO_CONFIG.slidesLogo || {};

  // Primary: URL substring match
  const substrings = cfg.oldContentUrlSubstrings || [];
  if (substrings.length > 0) {
    const contentUrl = element.image.contentUrl || "";
    const sourceUrl  = element.image.sourceUrl  || "";
    for (var i = 0; i < substrings.length; i++) {
      const s = substrings[i];
      if (!s) continue;
      if (contentUrl.indexOf(s) !== -1 || sourceUrl.indexOf(s) !== -1) {
        return { matchedBy: "contentUrl" };
      }
    }
  }

  // Fallback: zone + size/aspect
  const zones = cfg.zones || [];
  if (zones.length === 0) return null;

  const tx = element.transform.translateX || 0;
  const ty = element.transform.translateY || 0;
  const w  = element.size.width.magnitude;
  const h  = element.size.height.magnitude;

  const centerX   = (tx + w / 2) / pageWidth;
  const centerY   = (ty + h / 2) / pageHeight;
  const widthPct  = w / pageWidth;
  const heightPct = h / pageHeight;
  const aspect    = h > 0 ? w / h : 0;

  // Apply size/aspect filter to zone-fallback candidates.
  const sb = cfg.sizeBounds;
  if (sb) {
    if (widthPct  < sb.minWidthPct  || widthPct  > sb.maxWidthPct)  return null;
    if (heightPct < sb.minHeightPct || heightPct > sb.maxHeightPct) return null;
    if (aspect    < sb.minAspect    || aspect    > sb.maxAspect)    return null;
  }

  for (var j = 0; j < zones.length; j++) {
    const z = zones[j];
    if (centerX >= z.xMin && centerX <= z.xMax &&
        centerY >= z.yMin && centerY <= z.yMax) {
      return { matchedBy: "zone", zoneName: z.name };
    }
  }

  return null;
}

// ---------------------------------------------------------------------------
// Step 18 — collectLogoMatches
// ---------------------------------------------------------------------------

/**
 * Walks masters, layouts, and slides and returns descriptors for every
 * page element classified as a logo. Each descriptor records enough
 * information to reproduce the element at the same position and size if
 * a delete-and-recreate fallback is required.
 *
 * NOTE: Elements nested inside elementGroup are not recursed into.
 *
 * @param {Object[]} taggedPages  Array of { page, kind } where kind is
 *                                "master" | "layout" | "slide".
 * @param {number}   pageWidth    Slide width in EMUs.
 * @param {number}   pageHeight   Slide height in EMUs.
 * @returns {Object[]} Array of {
 *     objectId, parentPageId, pageKind, pageName,
 *     transform, size, matchedBy, zoneName
 *   }.
 */
function collectLogoMatches(taggedPages, pageWidth, pageHeight) {
  const matches = [];

  taggedPages.forEach(function(entry) {
    const page = entry.page;
    const pageName = page.pageProperties && page.pageProperties.name
      ? page.pageProperties.name
      : page.objectId;

    (page.pageElements || []).forEach(function(element) {
      const result = classifyLogoElement(element, pageWidth, pageHeight);
      if (!result) return;

      matches.push({
        objectId:     element.objectId,
        parentPageId: page.objectId,
        pageKind:     entry.kind,
        pageName:     pageName,
        transform:    element.transform,
        size:         element.size,
        matchedBy:    result.matchedBy,
        zoneName:     result.zoneName || null,
      });
    });
  });

  return matches;
}

// ---------------------------------------------------------------------------
// Step 19 — replaceLogos
// ---------------------------------------------------------------------------

/**
 * Replaces logo images on master, layout, AND slide pages using a layered
 * detection strategy (image-source URL substring, falling back to position
 * zones with a size/aspect filter).
 *
 * Each match is processed individually rather than as one large batch, so a
 * single failure does not poison the rest. When SlidesApp.Image.replace
 * fails (e.g. "can't replace a placeholder image"), the element is deleted
 * and a fresh image is created at the same transform and size via the REST
 * API. Recreated images are no longer placeholders, so subsequent runs
 * succeed via the normal replace path.
 *
 * Always run with dryRun=true first to audit matches before committing.
 *
 * Limitations:
 *   - Elements nested inside elementGroup are not traversed.
 *   - Hyperlinks and alt-text on the original element are lost when the
 *     element is recreated (the replace path preserves them).
 *
 * @param {string}  presentationId
 * @param {boolean} [dryRun=false]
 * @param {Object}  [cachedPresentation]  Pre-fetched presentation object; fetched if omitted.
 */
function replaceLogos(presentationId, dryRun, cachedPresentation) {
  const isDryRun = dryRun === true;
  const presentation = cachedPresentation || getPresentation(presentationId);
  const pageWidth  = presentation.pageSize.width.magnitude;
  const pageHeight = presentation.pageSize.height.magnitude;

  const taggedPages = [].concat(
    (presentation.masters || []).map(function(p) { return { page: p, kind: "master" }; }),
    (presentation.layouts || []).map(function(p) { return { page: p, kind: "layout" }; }),
    (presentation.slides  || []).map(function(p) { return { page: p, kind: "slide"  }; })
  );

  const matches = collectLogoMatches(taggedPages, pageWidth, pageHeight);

  if (matches.length === 0) {
    Logger.log("replaceLogos: no logo elements matched.");
    return;
  }

  if (isDryRun) {
    matches.forEach(function(m) {
      Logger.log(
        "[DRY RUN] Would replace logo: objectId=%s pageKind=%s page=%s matchedBy=%s zone=%s",
        m.objectId, m.pageKind, m.pageName, m.matchedBy, m.zoneName || "-"
      );
    });
    Logger.log("replaceLogos: %d match(es) found (dry run).", matches.length);
    return;
  }

  // Build an objectId → match descriptor index for quick lookup during
  // the SlidesApp pass.
  const matchById = {};
  matches.forEach(function(m) { matchById[m.objectId] = m; });

  // Fetch the logo blob once for the SlidesApp.Image.replace path.
  const logoBlob   = DriveApp.getFileById(LOGO_CONFIG.newLogoFileId).getBlob();
  const newLogoUrl = LOGO_CONFIG.newLogoUrl;
  const deck       = SlidesApp.openById(presentationId);

  // Track which matches succeeded via Image.replace so we can recreate the rest.
  const handled = {};
  var replacedCount  = 0;
  var recreatedCount = 0;
  var failedCount    = 0;

  function tryReplaceImages(images) {
    images.forEach(function(img) {
      const oid = img.getObjectId();
      if (!matchById[oid] || handled[oid]) return;
      try {
        img.replace(logoBlob, false);
        handled[oid] = "replaced";
        replacedCount++;
        Logger.log("replaceLogos: replaced objectId=%s", oid);
      } catch (err) {
        // Most commonly: "Can't replace a placeholder image."
        Logger.log("replaceLogos: replace failed for objectId=%s — %s. Will recreate.",
          oid, err && err.message ? err.message : err);
      }
    });
  }

  deck.getMasters().forEach(function(p) { tryReplaceImages(p.getImages()); });
  deck.getLayouts().forEach(function(p) { tryReplaceImages(p.getImages()); });
  deck.getSlides().forEach(function(p)  { tryReplaceImages(p.getImages()); });

  // Recreate any matches that Image.replace did not handle (placeholders or
  // elements not exposed via SlidesApp.getImages()).
  matches.forEach(function(m) {
    if (handled[m.objectId]) return;

    try {
      const requests = [
        { deleteObject: { objectId: m.objectId } },
        {
          createImage: {
            url: newLogoUrl,
            elementProperties: {
              pageObjectId: m.parentPageId,
              size:         m.size,
              transform:    m.transform,
            },
          },
        },
      ];
      Slides.Presentations.batchUpdate({ requests: requests }, presentationId);
      handled[m.objectId] = "recreated";
      recreatedCount++;
      Logger.log("replaceLogos: recreated objectId=%s on page=%s (matchedBy=%s zone=%s)",
        m.objectId, m.pageName, m.matchedBy, m.zoneName || "-");
    } catch (err) {
      handled[m.objectId] = "failed";
      failedCount++;
      Logger.log("replaceLogos: FAILED to recreate objectId=%s — %s",
        m.objectId, err && err.message ? err.message : err);
    }
  });

  Logger.log("replaceLogos: done. replaced=%d recreated=%d failed=%d (of %d matches)",
    replacedCount, recreatedCount, failedCount, matches.length);
}

// ---------------------------------------------------------------------------
// Step 10 — updateSlidesPresentation (public orchestrator)
// ---------------------------------------------------------------------------

/**
 * Runs the full brand update pipeline on a single presentation:
 *   1. Update master theme ColorScheme (Accent slots → new palette)
 *   2. Replace all inline (direct) RGB colors
 *   3. Replace Poppins / Figtree fonts with Lexend
 *   4. Replace logo images on master, layout, and slide pages
 *
 * @param {string}  presentationId
 * @param {boolean} [dryRun=false]  Passed through to replaceLogos.
 */
function updateSlidesPresentation(presentationId, dryRun) {
  const presentation = getPresentation(presentationId);

  Logger.log("Starting brand update for presentation: %s", presentationId);

  // 1. Theme color scheme (masters only)
  updateMasterThemeColors(presentationId, presentation.masters || []);
  Logger.log("  ✓ Master theme colors updated");

  // 2. Inline (direct) colors across all pages
  replaceInlineColors(presentationId, presentation);
  Logger.log("  ✓ Inline colors replaced");

  // 3. Fonts across all pages
  replaceFonts(presentationId, presentation);
  Logger.log("  ✓ Fonts replaced");

  // 4. Logos on master, layout, and slide pages
  replaceLogos(presentationId, dryRun, presentation);
  Logger.log("  ✓ Logo replacement %s", dryRun ? "dry run complete" : "complete");
}
