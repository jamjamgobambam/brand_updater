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

    // --- Page elements (recurse into grouped children) ---
    // Grouped shapes live under element.elementGroup.children and are
    // otherwise skipped entirely, so colors inside a group (e.g. decorative
    // ellipses/labels) never get recolored. Recurse so they're covered.
    function processElement(element) {
      if (element.elementGroup && element.elementGroup.children) {
        element.elementGroup.children.forEach(processElement);
      }
      const eid = element.objectId;

      // Shape fill
      // Skip placeholder shapes — updateShapeProperties is forbidden on placeholders
      // that live on a master or layout slide (Slides API restriction).
      // Skip if propertyState is "NOT_RENDERED" — the fill is explicitly hidden.
      // The API still returns solidFill.color for hidden fills (stored-but-not-rendered),
      // so without this guard we'd match that color and accidentally make the fill visible.
      const shapeFillRgb =
        element.shape &&
        !element.shape.placeholder &&
        element.shape.shapeProperties &&
        element.shape.shapeProperties.shapeBackgroundFill &&
        element.shape.shapeProperties.shapeBackgroundFill.propertyState !== "NOT_RENDERED" &&
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

      // Shape outline — same guards: skip placeholders and hidden outlines
      const outlineRgb =
        element.shape &&
        !element.shape.placeholder &&
        element.shape.shapeProperties &&
        element.shape.shapeProperties.outline &&
        element.shape.shapeProperties.outline.propertyState !== "NOT_RENDERED" &&
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
      // Placeholder shapes on masters/layouts are excluded here — updateTextStyle
      // is also forbidden on them. They are handled by replacePlaceholderColors.
      const textElements =
        element.shape &&
        !element.shape.placeholder &&
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
              cell.tableCellProperties.tableCellBackgroundFill.propertyState !== "NOT_RENDERED" &&
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
    }
    (page.pageElements || []).forEach(processElement);
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
    // Recurse into grouped children (element.elementGroup.children) so fonts
    // inside groups are replaced too — they are otherwise skipped entirely.
    function processElement(element) {
      if (element.elementGroup && element.elementGroup.children) {
        element.elementGroup.children.forEach(processElement);
      }
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
    }
    (page.pageElements || []).forEach(processElement);
  });

  return requests;
}

// ---------------------------------------------------------------------------
// Step 13 — replaceFonts
// ---------------------------------------------------------------------------

/**
 * Replaces all explicit Poppins/Figtree font references with Geist across
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
// Step 15 — logSectionStructure (diagnostic utility)
// ---------------------------------------------------------------------------

/**
 * Dumps the NATIVE structure of a deck — SLIDES, LAYOUTS, and MASTERS — so we
 * can copy the exact design of the section-divider and top-bar elements out of
 * a correctly-branded deck and reproduce it programmatically on decks where
 * those elements are baked-in raster images.
 *
 * CRITICAL: in these decks the section-divider splash design and the colored
 * top bar usually live on the LAYOUT (or master), not on the slide — a divider
 * slide is often empty at the slide level and just references a "section
 * header" layout. So this walks layouts and masters too, and reports which
 * layout each slide uses, so a blank divider slide can be mapped to its layout.
 *
 * For every page it logs, per element (recursing into GROUPS):
 *   - SHAPE: type, geometry (points AND % of slide), solid-fill hex (or theme
 *     color), and each text run's text / font / size / foreground color
 *   - IMAGE: geometry only (these are what we REPLACE on broken decks)
 *   - the page background fill (solid hex / theme color)
 *
 * Slides/layouts whose text contains a section word are flagged with >>>.
 *
 * Uses SlidesApp (not the REST API) so colors are already resolved to RGB and
 * geometry is reported in points — the same unit insertShape/insertTextBox use.
 *
 * @param {string} [presentationId]  Deck to inspect; defaults to the active
 *                                   presentation when omitted.
 */
function logSectionStructure(presentationId) {
  const deck = presentationId
    ? SlidesApp.openById(presentationId)
    : SlidesApp.getActivePresentation();
  const pageW = deck.getPageWidth();
  const pageH = deck.getPageHeight();
  const SECTION_WORDS = ["warm up", "activity", "wrap up"];

  Logger.log("=== SECTION STRUCTURE: %s (page %sx%s pt) ===",
    deck.getId(), pageW.toFixed(0), pageH.toFixed(0));

  function solidHex_(color) {
    // color: a SlidesApp Color (from fill/text). Returns "#rrggbb",
    // "theme:NAME", or null.
    try {
      if (!color) return null;
      if (color.getColorType() === SlidesApp.ColorType.RGB) {
        var c = color.asRgbColor();
        return ("#" +
          (c.getRed()   < 16 ? "0" : "") + c.getRed().toString(16) +
          (c.getGreen() < 16 ? "0" : "") + c.getGreen().toString(16) +
          (c.getBlue()  < 16 ? "0" : "") + c.getBlue().toString(16));
      }
      if (color.getColorType() === SlidesApp.ColorType.THEME) {
        return "theme:" + color.asThemeColor().getThemeColorType();
      }
    } catch (e) { /* fall through */ }
    return null;
  }

  function geom_(el) {
    var L = el.getLeft(), T = el.getTop(), W = el.getWidth(), H = el.getHeight();
    return Utilities.formatString(
      "x=%spt(%.0f%%) y=%spt(%.0f%%) w=%spt(%.0f%%) h=%spt(%.0f%%)",
      L.toFixed(1), 100 * L / pageW, T.toFixed(1), 100 * T / pageH,
      W.toFixed(1), 100 * W / pageW, H.toFixed(1), 100 * H / pageH);
  }

  // Dump a single PageElement, recursing into groups. `pad` is leading spaces.
  function dumpElement_(el, pad) {
    var type;
    try { type = el.getPageElementType(); } catch (e) { return; }

    if (type === SlidesApp.PageElementType.GROUP) {
      Logger.log("%sGROUP [%s] %s", pad, el.getObjectId(), geom_(el));
      el.asGroup().getChildren().forEach(function(child) {
        dumpElement_(child, pad + "  ");
      });
      return;
    }

    if (type === SlidesApp.PageElementType.IMAGE) {
      Logger.log("%sIMAGE [%s] %s", pad, el.getObjectId(), geom_(el));
      return;
    }

    if (type === SlidesApp.PageElementType.SHAPE) {
      var shape = el.asShape();
      var fillHex = null;
      try {
        var fill = shape.getFill();
        if (fill.getType() === SlidesApp.FillType.SOLID) {
          fillHex = solidHex_(fill.getSolidFill().getColor());
        }
      } catch (e) { /* placeholder fill can throw */ }

      var typeStr;
      try { typeStr = String(shape.getShapeType()); } catch (e) { typeStr = "?"; }

      Logger.log("%sSHAPE %s [%s] fill=%s %s",
        pad, typeStr, shape.getObjectId(), fillHex, geom_(shape));

      try {
        shape.getText().getRuns().forEach(function(run) {
          var txt = run.asString().replace(/\n/g, "\\n");
          if (!txt.trim()) return;
          var st = run.getTextStyle();
          Logger.log("%s  text:\"%s\" font=%s size=%s color=%s",
            pad, txt.length > 60 ? txt.substring(0, 60) + "…" : txt,
            st.getFontFamily(), st.getFontSize(), solidHex_(st.getForegroundColor()));
        });
      } catch (e) { /* no text */ }
      return;
    }

    // LINE, TABLE, VIDEO, WORD_ART, SHEETS_CHART, etc.
    Logger.log("%s%s [%s] %s", pad, String(type), el.getObjectId(), geom_(el));
  }

  // Dump a whole page (slide / layout / master). Returns true if its text
  // contains a section word.
  function dumpPage_(page, label) {
    var els = page.getPageElements();
    var allText = els.map(function(e) {
      try {
        return e.getPageElementType() === SlidesApp.PageElementType.SHAPE
          ? e.asShape().getText().asString() : "";
      } catch (err) { return ""; }
    }).join(" ").toLowerCase();
    var isSection = SECTION_WORDS.some(function(w) { return allText.indexOf(w) !== -1; });

    Logger.log("\n%s %s — %d element(s)", isSection ? ">>>" : "   ", label, els.length);

    try {
      var bg = page.getBackground();
      var bt = bg.getType();
      if (bt === SlidesApp.PageBackgroundType.SOLID) {
        Logger.log("    background: SOLID %s", solidHex_(bg.getSolidFill().getColor()));
      } else if (bt === SlidesApp.PageBackgroundType.PICTURE) {
        // PICTURE backgrounds are baked-in raster art — the color/font passes
        // cannot touch them. This is what we need to convert to native.
        var pf = bg.getPictureFill && bg.getPictureFill();
        var url = "(no url)";
        try { url = pf ? pf.getContentUrl() : "(no fill)"; } catch (e2) {}
        Logger.log("    background: PICTURE %s", url);
      } else {
        // NONE = inherits from layout/master; UNSUPPORTED = something else.
        Logger.log("    background: %s", String(bt));
      }
    } catch (e) {
      Logger.log("    background: (read failed: %s)", e && e.message);
    }

    els.forEach(function(el) { dumpElement_(el, "    "); });
    return isSection;
  }

  // --- Slides (report the layout each one uses) ---
  Logger.log("\n########## SLIDES ##########");
  deck.getSlides().forEach(function(slide, i) {
    var layName = "?";
    try {
      var lay = slide.getLayout();
      layName = lay ? (lay.getLayoutName() + " / " + lay.getObjectId()) : "(none)";
    } catch (e) { /* layout unavailable */ }
    dumpPage_(slide, "slide[" + i + "]  layout=" + layName);
  });

  // --- Layouts (where divider/top-bar designs usually live) ---
  Logger.log("\n########## LAYOUTS ##########");
  deck.getLayouts().forEach(function(layout, i) {
    var name = "?";
    try { name = layout.getLayoutName(); } catch (e) {}
    dumpPage_(layout, "layout[" + i + "]  name=" + name + " / " + layout.getObjectId());
  });

  // --- Masters ---
  Logger.log("\n########## MASTERS ##########");
  deck.getMasters().forEach(function(master, i) {
    dumpPage_(master, "master[" + i + "]  " + master.getObjectId());
  });

  Logger.log("\n=== DONE ===");
}

// ---------------------------------------------------------------------------
// Step 17 — Gemini-based logo replacement
// ---------------------------------------------------------------------------

/**
 * Walks every image on a Slides page (including images nested in groups) and
 * applies the supplied callback. Tables in Slides cannot contain images so
 * they are intentionally not traversed here.
 *
 * @param {SlidesApp.Page} page
 * @param {(image: SlidesApp.Image, page: SlidesApp.Page) => void} fn
 */
function forEachSlidesImage_(page, fn) {
  page.getImages().forEach(function(img) { fn(img, page); });
  page.getGroups().forEach(function(group) { walkSlidesGroupImages_(group, page, fn); });
}

function walkSlidesGroupImages_(group, page, fn) {
  group.getChildren().forEach(function(child) {
    var t = child.getPageElementType();
    if (t === SlidesApp.PageElementType.IMAGE) {
      fn(child.asImage(), page);
    } else if (t === SlidesApp.PageElementType.GROUP) {
      walkSlidesGroupImages_(child.asGroup(), page, fn);
    }
  });
}

/**
 * Replace a single image element with the new CodeAI logo.
 *
 * The new logo is sized by TARGET WIDTH (height follows the new logo's
 * natural aspect ratio), chosen by the original's position on the slide:
 *
 *   - Top-left header band    → width = max(origWidth, HEADER_LOGO_MIN_WIDTH_PT)
 *                                and adjacent text shapes are shifted right
 *                                so they don't overlap the (square-ish)
 *                                replacement logo.
 *   - Top-center title band   → width = origWidth * TITLE_LOGO_WIDTH_SCALE
 *   - Bottom-right corner band → width = max(origWidth, CORNER_LOGO_MIN_WIDTH_PT)
 *   - Anywhere else            → width = origWidth * LOGO_SCALE
 *
 * The replacement is centered on the original's center, then clamped so the
 * new logo never crosses the slide edge (LOGO_SLIDE_MARGIN of padding).
 *
 * Sizing by width (not by a uniform W×H scale of the original box) matters
 * because the legacy wordmark was wide-and-short while the new logo is
 * roughly square — scaling by height made the replacement look tiny.
 *
 * @param {SlidesApp.Image} image  The image to replace.
 * @param {SlidesApp.Page} page    The page the image belongs to.
 * @param {number} pageWidth       Slide width in points.
 * @param {number} pageHeight      Slide height in points.
 * @param {SlidesApp.Page[]} [downstreamPages]
 *        Pages whose shapes inherit from `page` (e.g. layouts of a master,
 *        or slides of a layout). Used by the header-band shift to reach
 *        slide-level text boxes when the logo lives on the master/layout.
 */
function replaceSlidesLogoImage_(image, page, pageWidth, pageHeight, downstreamPages) {
  const left = image.getLeft();
  const top  = image.getTop();
  const w    = image.getWidth();
  const h    = image.getHeight();
  const rotation = image.getRotation ? image.getRotation() : 0;

  // Original center as a fraction of slide dimensions, used to pick a
  // position-aware target width.
  const centerX = pageWidth  ? (left + w / 2) / pageWidth  : 0.5;
  const centerY = pageHeight ? (top  + h / 2) / pageHeight : 0.5;

  const titleCfg  = LOGO_CONFIG.titleLogo;
  const cornerCfg = LOGO_CONFIG.cornerLogo;
  const headerCfg = LOGO_CONFIG.headerLogo;
  const inHeaderBand = !!headerCfg &&
    centerX <= headerCfg.xMax &&
    centerY <= headerCfg.yMax;
  const inTitleBand = !inHeaderBand &&
    centerX >= titleCfg.xMin &&
    centerX <= titleCfg.xMax &&
    centerY <= 0.45; // a touch more permissive than titleLogo.yMax for tall titles
  const inCornerBand = !inHeaderBand &&
    centerX >= cornerCfg.xThreshold &&
    centerY >= cornerCfg.yThreshold;

  // Insert at default size first so we can read the new logo's natural
  // aspect ratio, then size & position correctly.
  const newBlob  = getCodeAILogoBlob_();
  const inserted = page.insertImage(newBlob);
  const naturalW = inserted.getInherentWidth();
  const naturalH = inserted.getInherentHeight();
  const aspect = (naturalW && naturalH) ? naturalW / naturalH : (w / h || 1);

  // Pick target width based on position, then derive height from aspect.
  let targetW;
  if (inHeaderBand) {
    targetW = Math.max(w, HEADER_LOGO_MIN_WIDTH_PT);
  } else if (inTitleBand) {
    targetW = w * TITLE_LOGO_WIDTH_SCALE;
  } else if (inCornerBand) {
    targetW = Math.max(w, CORNER_LOGO_MIN_WIDTH_PT);
  } else {
    targetW = w * LOGO_SCALE;
  }

  // Cap width so the logo (with margin on each side) always fits the slide.
  if (pageWidth) {
    const maxAllowedW = pageWidth - 2 * LOGO_SLIDE_MARGIN;
    if (maxAllowedW > 0 && targetW > maxAllowedW) targetW = maxAllowedW;
  }

  let finalW = targetW;
  let finalH = finalW / aspect;

  // If height now overflows the slide, scale both back down proportionally.
  if (pageHeight) {
    const maxAllowedH = pageHeight - 2 * LOGO_SLIDE_MARGIN;
    if (maxAllowedH > 0 && finalH > maxAllowedH) {
      finalH = maxAllowedH;
      finalW = finalH * aspect;
    }
  }

  // Center on the original's center.
  let finalLeft = left + (w - finalW) / 2;
  let finalTop  = top  + (h - finalH) / 2;

  // Clamp to slide bounds with a margin so an edge-anchored original (e.g.
  // a corner logo) doesn't push the enlarged replacement off the slide.
  if (pageWidth) {
    const minLeft = LOGO_SLIDE_MARGIN;
    const maxLeft = pageWidth - finalW - LOGO_SLIDE_MARGIN;
    finalLeft = (maxLeft < minLeft)
      ? (pageWidth - finalW) / 2
      : Math.max(minLeft, Math.min(maxLeft, finalLeft));
  }
  if (pageHeight) {
    const minTop = LOGO_SLIDE_MARGIN;
    const maxTop = pageHeight - finalH - LOGO_SLIDE_MARGIN;
    finalTop = (maxTop < minTop)
      ? (pageHeight - finalH) / 2
      : Math.max(minTop, Math.min(maxTop, finalTop));
  }

  inserted.setLeft(finalLeft);
  inserted.setTop(finalTop);
  inserted.setWidth(finalW);
  inserted.setHeight(finalH);
  if (rotation && inserted.setRotation) {
    try { inserted.setRotation(rotation); } catch (e) { /* not all elements support rotation */ }
  }

  // Shift adjacent text shapes out from under the new top-left header logo.
  // Done before image.remove() so the inserted logo's id is known and skipped.
  if (inHeaderBand) {
    shiftOverlappingTextShapes_(
      page,
      inserted.getObjectId(),
      { left: finalLeft, top: finalTop, right: finalLeft + finalW, bottom: finalTop + finalH },
      pageWidth,
      downstreamPages || []
    );
  }

  image.remove();
}

/**
 * After replacing a top-left header logo, shift any text shape on the same
 * page (and any downstream pages — layouts of a master, slides of a layout)
 * whose bounding box overlaps the new logo (or another already-shifted
 * shape) to the right of that obstacle's right edge plus HEADER_TEXT_GAP_PT.
 *
 * Cascading: each successful shift becomes a new obstacle, so a small CSF
 * tag on the master that gets pushed right will in turn push a Course-A
 * box on the slide.
 *
 * Width is left unchanged so the text reflows; if shifting would push the
 * shape past the right edge of the slide (minus LOGO_SLIDE_MARGIN), the
 * shape is left in place and a warning is logged.
 *
 * Recurses into groups so nested header text shapes are also handled.
 *
 * @param {SlidesApp.Page} page
 * @param {string} skipObjectId            Object id of the just-inserted logo.
 * @param {{left:number,top:number,right:number,bottom:number}} logoRect
 * @param {number} pageWidth               Slide width in points.
 * @param {SlidesApp.Page[]} [downstreamPages]
 */
function shiftOverlappingTextShapes_(page, skipObjectId, logoRect, pageWidth, downstreamPages) {
  // Collect every candidate shape (top-level + nested in groups) once, with
  // its current geometry. We run multiple passes so a shape that gets shifted
  // out of the logo's path becomes a new obstacle for shapes further to the
  // right (e.g. the small CSF tag pushes the long course-name box over too).
  var candidates = [];
  var seenIds = {};

  // A shape only counts as a header neighbor if its vertical center lies
  // inside the logo bar's Y range. Without this, the relaxed rectsOverlap
  // would later match large content-area boxes (e.g. a body text frame that
  // spans most of the slide) against the small logo via the obstacle's
  // center, and we'd try to shift them — they always overflow.
  function inLogoYBand(top, height) {
    var midY = top + height / 2;
    return midY >= logoRect.top && midY <= logoRect.bottom;
  }

  function collectShape(shape) {
    if (!shape) return;
    var id = shape.getObjectId && shape.getObjectId();
    if (id === skipObjectId) return;
    if (id && seenIds[id]) return;
    var sLeft, sTop, sW, sH;
    try {
      sLeft = shape.getLeft();
      sTop  = shape.getTop();
      sW    = shape.getWidth();
      sH    = shape.getHeight();
    } catch (e) {
      return; // some placeholders refuse geometry reads on master/layout
    }
    if (!inLogoYBand(sTop, sH)) return;
    if (id) seenIds[id] = true;
    candidates.push({
      shape: shape,
      left: sLeft, top: sTop, width: sW, height: sH,
      shifted: false,
    });
  }

  function walkGroup(group) {
    group.getChildren().forEach(function(child) {
      var t = child.getPageElementType();
      if (t === SlidesApp.PageElementType.SHAPE) {
        collectShape(child.asShape());
      } else if (t === SlidesApp.PageElementType.GROUP) {
        walkGroup(child.asGroup());
      }
    });
  }

  function collectFromPage(p) {
    p.getShapes().forEach(collectShape);
    p.getGroups().forEach(walkGroup);
  }

  collectFromPage(page);
  (downstreamPages || []).forEach(collectFromPage);

  // Obstacle list — starts with the new logo, grows with every successful
  // shift so subsequent passes can detect cascading overlaps.
  var obstacles = [{
    left: logoRect.left, top: logoRect.top,
    right: logoRect.right, bottom: logoRect.bottom,
  }];

  function rectsOverlap(a, b) {
    // Strict horizontal AABB. Vertical: relaxed — accept if either rect's
    // vertical center lies inside the other's Y range. Header text shapes
    // often have slightly different top/height than the logo bar (line-height
    // padding, baseline offsets), and slide-level text boxes can sit at a
    // different Y from the master logo, so strict AABB Y misses them.
    if (!(a.left < b.right && a.right > b.left)) return false;
    var aMid = (a.top + a.bottom) / 2;
    var bMid = (b.top + b.bottom) / 2;
    return (aMid >= b.top && aMid <= b.bottom) ||
           (bMid >= a.top && bMid <= a.bottom);
  }

  // Iterate until no candidate gets shifted in a full pass. Cap at a few
  // passes to avoid runaway loops on pathological inputs.
  for (var pass = 0; pass < 8; pass++) {
    var anyShifted = false;

    for (var i = 0; i < candidates.length; i++) {
      var c = candidates[i];
      if (c.shifted) continue;

      var cRect = {
        left: c.left, top: c.top,
        right: c.left + c.width, bottom: c.top + c.height,
      };

      // Find the rightmost obstacle this shape overlaps.
      var hitRight = -Infinity;
      var fullBleed = false;
      for (var k = 0; k < obstacles.length; k++) {
        var o = obstacles[k];
        if (rectsOverlap(cRect, o)) {
          if (o.right > hitRight) hitRight = o.right;
          // Treat as full-bleed background ONLY if the shape extends past
          // the obstacle on BOTH sides (i.e. envelops it). A shape that
          // merely starts before the obstacle's left edge but ends inside
          // it is normal adjacent text and should be shifted.
          if (cRect.left <= o.left && cRect.right >= o.right) {
            fullBleed = true;
          }
        }
      }
      if (hitRight === -Infinity) continue; // no overlap
      if (fullBleed) continue;               // skip background bars

      var newLeft = hitRight + HEADER_TEXT_GAP_PT;
      if (newLeft <= c.left) continue; // already clear

      if (pageWidth && newLeft + c.width > pageWidth - LOGO_SLIDE_MARGIN) {
        Logger.log(
          "Header shift skipped — shape %s would overflow slide (newLeft=%s width=%s pageWidth=%s)",
          c.shape.getObjectId ? c.shape.getObjectId() : "?",
          newLeft.toFixed(1), c.width.toFixed(1), pageWidth.toFixed(1)
        );
        c.shifted = true; // don't retry
        continue;
      }

      try {
        c.shape.setLeft(newLeft);
      } catch (e) {
        Logger.log(
          "Header shift failed for shape %s: %s",
          c.shape.getObjectId ? c.shape.getObjectId() : "?", e && e.message
        );
        c.shifted = true;
        continue;
      }

      c.left = newLeft;
      c.shifted = true;
      anyShifted = true;
      obstacles.push({
        left: newLeft, top: c.top,
        right: newLeft + c.width, bottom: c.top + c.height,
      });
    }

    if (!anyShifted) break;
  }
}

// ---------------------------------------------------------------------------
// Step 19 — replaceLogos
// ---------------------------------------------------------------------------

/**
 * Replaces logo images across masters, layouts, and slides using a Gemini-
 * based image classifier. Every image is sent to Gemini (cached by SHA-256
 * so identical images cost one call across the run). Images classified as a
 * standalone code.org logo with confidence ≥ LOGO_CONFIDENCE.REPLACE are
 * replaced; results in the [REVIEW, REPLACE) band are logged for manual
 * review and left alone.
 *
 * Requires the GEMINI_API_KEY Script Property to be set; without it, every
 * image is skipped and a warning is logged.
 *
 * @param {string}  presentationId
 * @param {boolean} [dryRun=false]  When true, classifies & logs but makes no edits.
 * @param {Object}  [_cachedPresentation]  Unused — accepted for backwards
 *                                         compatibility with prior callers.
 */
function replaceLogos(presentationId, dryRun, _cachedPresentation) {
  const isDryRun = dryRun === true;
  const deck = SlidesApp.openById(presentationId);
  const pageWidth  = deck.getPageWidth();
  const pageHeight = deck.getPageHeight();

  // Counts per disposition so the user can see the classifier outcome.
  const counts = { replaced: 0, reviewed: 0, skipped: 0 };
  const reviewLog = [];

  function visit(page, pageLabel, downstreamPages) {
    forEachSlidesImage_(page, function(image) {
      let blob;
      try {
        blob = image.getBlob();
      } catch (e) {
        counts.skipped++;
        return;
      }

      const verdict = classifyLogo_(blob);
      const action  = logoAction_(verdict);

      if (action === "skip") {
        counts.skipped++;
        return;
      }
      if (action === "review") {
        counts.reviewed++;
        reviewLog.push(
          "  ? review (" + verdict.confidence.toFixed(2) + "): " +
          pageLabel + " / " + image.getObjectId()
        );
        return;
      }

      // action === "replace"
      if (isDryRun) {
        counts.replaced++;
        Logger.log(
          "[DRY RUN] Would replace logo (%.2f): %s / %s",
          verdict.confidence, pageLabel, image.getObjectId()
        );
        return;
      }

      try {
        replaceSlidesLogoImage_(image, page, pageWidth, pageHeight, downstreamPages);
        counts.replaced++;
      } catch (err) {
        counts.skipped++;
        Logger.log(
          "Logo replacement failed on %s / %s: %s",
          pageLabel, image.getObjectId(), err && err.message
        );
      }
    });
  }

  // Pre-compute "downstream" pages for each master/layout so the header-band
  // shift can reach slide-level text boxes (e.g. CSF on master, Course-A on
  // slide). Logo on master  → downstream = its layouts + slides using those
  // layouts. Logo on layout → downstream = slides using that layout.
  const allSlides = deck.getSlides();

  deck.getMasters().forEach(function(m, i) {
    const masterLayouts = m.getLayouts();
    const masterLayoutIds = {};
    masterLayouts.forEach(function(l) { masterLayoutIds[l.getObjectId()] = true; });
    const slidesUnderMaster = allSlides.filter(function(s) {
      var lyt = s.getLayout && s.getLayout();
      return lyt && masterLayoutIds[lyt.getObjectId()];
    });

    visit(m, "master[" + i + "]", masterLayouts.concat(slidesUnderMaster));

    masterLayouts.forEach(function(l, j) {
      const layoutId = l.getObjectId();
      const slidesForLayout = allSlides.filter(function(s) {
        var lyt = s.getLayout && s.getLayout();
        return lyt && lyt.getObjectId() === layoutId;
      });
      visit(l, "master[" + i + "].layout[" + j + "]", slidesForLayout);
    });
  });
  allSlides.forEach(function(s, i) { visit(s, "slide[" + i + "]", []); });

  if (reviewLog.length) {
    Logger.log("Logo classifier — needs review:\n%s", reviewLog.join("\n"));
  }
  Logger.log(
    "Logo replacement %s — replaced:%d review:%d skipped:%d",
    isDryRun ? "dry run" : "complete",
    counts.replaced, counts.reviewed, counts.skipped
  );
}

// ---------------------------------------------------------------------------
// replacePlaceholderColors — SlidesApp fallback for master/layout placeholders
// ---------------------------------------------------------------------------

/**
 * Updates fill and border colors on placeholder shapes that live on master and
 * layout slides. The REST API updateShapeProperties request is forbidden on
 * placeholder shapes in masters/layouts, so SlidesApp is used instead.
 *
 * Non-placeholder shapes on masters/layouts are handled by replaceInlineColors
 * via the REST API batch.
 *
 * @param {string} presentationId
 */
function replacePlaceholderColors(presentationId) {
  var deck = SlidesApp.openById(presentationId);
  var pages = [];
  deck.getMasters().forEach(function(m) { pages.push(m); });
  deck.getLayouts().forEach(function(l) { pages.push(l); });

  pages.forEach(function(page) {
    page.getShapes().forEach(function(shape) {
      if (shape.getPlaceholderType() === SlidesApp.PlaceholderType.NONE) return;

      // Fill
      var fill = shape.getFill();
      if (fill.getType() === SlidesApp.FillType.SOLID) {
        var sc = fill.getSolidFill().getColor().asRgbColor();
        var fillRgb = { red: sc.getRed() / 255, green: sc.getGreen() / 255, blue: sc.getBlue() / 255 };
        var newFillHex = findColorMapping(fillRgb, COLOR_MAP, COLOR_DISTANCE_THRESHOLD);
        if (newFillHex) fill.setSolidFill(newFillHex);
      }

      // Border
      var border = shape.getBorder();
      if (border.isVisible()) {
        var lineFill = border.getLineFill();
        if (lineFill.getType() === SlidesApp.LineFillType.SOLID) {
          var bc = lineFill.getSolidFill().getColor().asRgbColor();
          var borderRgb = { red: bc.getRed() / 255, green: bc.getGreen() / 255, blue: bc.getBlue() / 255 };
          var newBorderHex = findColorMapping(borderRgb, COLOR_MAP, COLOR_DISTANCE_THRESHOLD);
          if (newBorderHex) lineFill.setSolidFill(newBorderHex);
        }
      }

      // Text run foreground colors
      shape.getText().getRuns().forEach(function(run) {
        var style = run.getTextStyle();
        var color = style.getForegroundColor();
        if (!color || color.getColorType() !== SlidesApp.ColorType.RGB) return;
        var rgb = color.asRgbColor();
        var textRgb = { red: rgb.getRed() / 255, green: rgb.getGreen() / 255, blue: rgb.getBlue() / 255 };
        var newTextHex = findColorMapping(textRgb, COLOR_MAP, COLOR_DISTANCE_THRESHOLD);
        if (newTextHex) style.setForegroundColor(newTextHex);
      });
    });
  });
}

// ---------------------------------------------------------------------------
// Step 10 — updateSlidesPresentation (public orchestrator)
// ---------------------------------------------------------------------------

/**
 * Runs the full brand update pipeline on a single presentation:
 *   1. Update master theme ColorScheme (Accent slots → new palette)
 *   2. Replace all inline (direct) RGB colors
 *   3. Replace Poppins / Figtree fonts with Geist
 *   4. Replace logo images on master/layout slides
 *
 * @param {string}  presentationId
 * @param {boolean} [dryRun=false]  Passed through to replaceLogos.
 */
function updateSlidesPresentation(presentationId, dryRun, options) {
  var opts = options || { colors: true, fonts: true, logo: true };
  const presentation = getPresentation(presentationId);

  Logger.log("Starting brand update for presentation: %s", presentationId);

  if (opts.colors) {
    updateMasterThemeColors(presentationId, presentation.masters || []);
    Logger.log("  ✓ Master theme colors updated");

    replaceInlineColors(presentationId, presentation);
    Logger.log("  ✓ Inline colors replaced");

    replacePlaceholderColors(presentationId);
    Logger.log("  ✓ Placeholder shape colors replaced");
  }

  if (opts.fonts) {
    replaceFonts(presentationId, presentation);
    Logger.log("  ✓ Fonts replaced");
  }

  if (opts.logo) {
    replaceLogos(presentationId, dryRun, presentation);
    Logger.log("  ✓ Logo replacement %s", dryRun ? "dry run complete" : "complete");
  }
}
