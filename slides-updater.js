// =============================================================================
// slides-updater.js — Google Slides brand updater (colors, fonts, logos)
// Depends on globals defined in utils.js: COLOR_MAP, FONT_MAP, LOGO_CONFIG,
// hexToNormalizedRgb, normalizedRgbMatches, driveFileUrl, HYPERLINK_NEW_HEX
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

  // Map Accent slot type → new hex using COLOR_MAP (positional: ACCENT1 first)
  const accentTypes = ["ACCENT1", "ACCENT2", "ACCENT3", "ACCENT4", "ACCENT5", "ACCENT6"];
  const accentNewHexes = COLOR_MAP.map(function(entry) { return entry.newHex; });

  masters.forEach(function(master) {
    const existingColors =
      master.pageProperties &&
      master.pageProperties.colorScheme &&
      master.pageProperties.colorScheme.colors;

    if (!existingColors) return;

    // Deep-copy then patch only the target slots
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
      colorMap.forEach(function(mapping) {
        if (normalizedRgbMatches(bgRgb, mapping.oldHex)) {
          requests.push({
            updatePageProperties: {
              objectId: pageId,
              pageProperties: {
                pageBackgroundFill: {
                  solidFill: {
                    color: { rgbColor: hexToNormalizedRgb(mapping.newHex) },
                  },
                },
              },
              fields: "pageBackgroundFill.solidFill.color",
            },
          });
        }
      });
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
        colorMap.forEach(function(mapping) {
          if (normalizedRgbMatches(shapeFillRgb, mapping.oldHex)) {
            requests.push({
              updateShapeProperties: {
                objectId: eid,
                shapeProperties: {
                  shapeBackgroundFill: {
                    solidFill: {
                      color: { rgbColor: hexToNormalizedRgb(mapping.newHex) },
                    },
                  },
                },
                fields: "shapeBackgroundFill.solidFill.color",
              },
            });
          }
        });
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
        colorMap.forEach(function(mapping) {
          if (normalizedRgbMatches(outlineRgb, mapping.oldHex)) {
            requests.push({
              updateShapeProperties: {
                objectId: eid,
                shapeProperties: {
                  outline: {
                    outlineFill: {
                      solidFill: {
                        color: { rgbColor: hexToNormalizedRgb(mapping.newHex) },
                      },
                    },
                  },
                },
                fields: "outline.outlineFill.solidFill.color",
              },
            });
          }
        });
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
          colorMap.forEach(function(mapping) {
            if (normalizedRgbMatches(fgRgb, mapping.oldHex)) {
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
                        rgbColor: hexToNormalizedRgb(mapping.newHex),
                      },
                    },
                  },
                  fields: "foregroundColor",
                },
              });
            }
          });
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

            if (!cellRgb) return;
            colorMap.forEach(function(mapping) {
              if (normalizedRgbMatches(cellRgb, mapping.oldHex)) {
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
                          color: { rgbColor: hexToNormalizedRgb(mapping.newHex) },
                        },
                      },
                    },
                    fields: "tableCellBackgroundFill.solidFill.color",
                  },
                });
              }
            });
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
        colorMap.forEach(function(mapping) {
          if (normalizedRgbMatches(lineRgb, mapping.oldHex)) {
            requests.push({
              updateLineProperties: {
                objectId: eid,
                lineProperties: {
                  lineFill: {
                    solidFill: {
                      color: { rgbColor: hexToNormalizedRgb(mapping.newHex) },
                    },
                  },
                },
                fields: "lineFill.solidFill.color",
              },
            });
          }
        });
      }
    });
  });

  return requests;
}

// ---------------------------------------------------------------------------
// Step 9 — replaceInlineColors
// ---------------------------------------------------------------------------

/**
 * Fetches the full presentation JSON and replaces all inline RGB colors
 * matching COLOR_MAP across masters, layouts, and slides.
 * Splits requests into batches of 400 to stay under the 500-request API limit.
 *
 * @param {string} presentationId
 */
function replaceInlineColors(presentationId) {
  const presentation = Slides.Presentations.get(presentationId);
  const allPages = [].concat(
    presentation.masters  || [],
    presentation.layouts  || [],
    presentation.slides   || []
  );

  const requests = buildInlineColorRequests(allPages, COLOR_MAP);
  if (requests.length === 0) return;

  // Batch in chunks of 400 to avoid hitting the ~500 request limit
  const BATCH_SIZE = 400;
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
      const textElements =
        element.shape &&
        element.shape.text &&
        element.shape.text.textElements;

      if (!textElements) return;

      textElements.forEach(function(te) {
        if (!te.textRun) return;
        const style = te.textRun.style || {};

        // Prefer weightedFontFamily (includes weight), fall back to fontFamily
        const wff = style.weightedFontFamily;
        const fontFamily = wff ? wff.fontFamily : style.fontFamily;
        if (!fontFamily) return; // null means inheriting — leave untouched

        fontMap.forEach(function(mapping) {
          if (fontFamily === mapping.oldFont) {
            const existingWeight = wff ? wff.weight : 400;
            requests.push({
              updateTextStyle: {
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
              },
            });
          }
        });
      });
    });
  });

  return requests;
}

// ---------------------------------------------------------------------------
// Step 13 — replaceFonts
// ---------------------------------------------------------------------------

/**
 * Fetches the full presentation JSON and replaces all explicit Poppins/Figtree
 * font references with Lexend across masters, layouts, and slides.
 * Splits requests into batches of 400.
 *
 * @param {string} presentationId
 */
function replaceFonts(presentationId) {
  const presentation = Slides.Presentations.get(presentationId);
  const allPages = [].concat(
    presentation.masters  || [],
    presentation.layouts  || [],
    presentation.slides   || []
  );

  const requests = buildFontRequests(allPages, FONT_MAP);
  if (requests.length === 0) return;

  const BATCH_SIZE = 400;
  for (let i = 0; i < requests.length; i += BATCH_SIZE) {
    const chunk = requests.slice(i, i + BATCH_SIZE);
    Slides.Presentations.batchUpdate({ requests: chunk }, presentationId);
  }
}

// ---------------------------------------------------------------------------
// Step 14 — logAllImages (diagnostic utility)
// ---------------------------------------------------------------------------

/**
 * Logs details of every image element on master and layout slides.
 * Run once on a representative presentation to determine the correct
 * position thresholds for LOGO_CONFIG before running replaceLogos.
 *
 * @param {string} presentationId
 */
function logAllImages(presentationId) {
  const presentation = Slides.Presentations.get(presentationId);
  const pageWidth  = presentation.pageSize.width.magnitude;
  const pageHeight = presentation.pageSize.height.magnitude;

  const pages = [].concat(
    presentation.masters || [],
    presentation.layouts || []
  );

  pages.forEach(function(page) {
    const pageName = page.pageProperties && page.pageProperties.name
      ? page.pageProperties.name
      : page.objectId;

    (page.pageElements || []).forEach(function(element) {
      if (!element.image) return;
      if (!element.transform) {
        Logger.log("Image [%s] on page [%s]: no transform (at default position)", element.objectId, pageName);
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

      Logger.log(
        "Image [%s] on page [%s]: centerX=%.3f centerY=%.3f width=%.3f height=%.3f sourceUrl=%s",
        element.objectId,
        pageName,
        centerX,
        centerY,
        widthPct,
        heightPct,
        element.image.sourceUrl || "(none)"
      );
    });
  });
}

// ---------------------------------------------------------------------------
// Step 17 — isLogoElement
// ---------------------------------------------------------------------------

/**
 * Returns true if an element is an image whose center falls within the
 * detection zone for the given logo type.
 *
 * @param {Object} element    A page element from the Slides API.
 * @param {number} pageWidth  Slide width in EMUs.
 * @param {number} pageHeight Slide height in EMUs.
 * @param {"corner"|"title"} type  Which logo zone to check.
 * @returns {boolean}
 */
function isLogoElement(element, pageWidth, pageHeight, type) {
  if (!element.image)     return false;
  if (!element.transform) return false;

  const tx = element.transform.translateX || 0;
  const ty = element.transform.translateY || 0;
  const w  = element.size.width.magnitude;
  const h  = element.size.height.magnitude;

  const centerX = (tx + w / 2) / pageWidth;
  const centerY = (ty + h / 2) / pageHeight;

  if (type === "corner") {
    return centerX > LOGO_CONFIG.cornerLogo.xThreshold &&
           centerY > LOGO_CONFIG.cornerLogo.yThreshold;
  }
  if (type === "title") {
    return centerX > LOGO_CONFIG.titleLogo.xMin  &&
           centerX < LOGO_CONFIG.titleLogo.xMax  &&
           centerY < LOGO_CONFIG.titleLogo.yMax;
  }
  return false;
}

// ---------------------------------------------------------------------------
// Step 18 — buildLogoReplaceRequests
// ---------------------------------------------------------------------------

/**
 * Iterates master and layout pages and builds replaceImage request objects
 * for every element identified as a corner or title logo.
 * In dry-run mode, logs matches instead of building requests.
 *
 * NOTE: Elements nested inside elementGroup are not recursed into.
 * Use logAllImages() to verify logo positions if nothing is matched.
 *
 * @param {Object[]} pages       Masters and layouts only.
 * @param {number}   pageWidth   Slide width in EMUs.
 * @param {number}   pageHeight  Slide height in EMUs.
 * @param {string}   newLogoUrl  Publicly accessible image URL for the new logo.
 * @param {boolean}  dryRun      When true, logs matches but builds no requests.
 * @returns {Object[]}           Array of replaceImage request objects.
 */
function buildLogoReplaceRequests(pages, pageWidth, pageHeight, newLogoUrl, dryRun) {
  const requests = [];

  pages.forEach(function(page) {
    const pageName = page.pageProperties && page.pageProperties.name
      ? page.pageProperties.name
      : page.objectId;

    (page.pageElements || []).forEach(function(element) {
      ["corner", "title"].forEach(function(logoType) {
        if (!isLogoElement(element, pageWidth, pageHeight, logoType)) return;

        if (dryRun) {
          Logger.log(
            "[DRY RUN] Would replace %s logo: objectId=%s page=%s",
            logoType,
            element.objectId,
            pageName
          );
        } else {
          requests.push({
            replaceImage: {
              imageObjectId: element.objectId,
              imageReplaceMethod: "CENTER_INSIDE",
              url: newLogoUrl,
            },
          });
        }
      });
    });
  });

  return requests;
}

// ---------------------------------------------------------------------------
// Step 19 — replaceLogos
// ---------------------------------------------------------------------------

/**
 * Replaces logo images on master and layout slides using position heuristics.
 * Always run with dryRun=true first to audit matches before committing.
 *
 * @param {string}  presentationId
 * @param {boolean} [dryRun=false]
 */
function replaceLogos(presentationId, dryRun) {
  const isDryRun = dryRun === true;
  const presentation = Slides.Presentations.get(presentationId);
  const pageWidth  = presentation.pageSize.width.magnitude;
  const pageHeight = presentation.pageSize.height.magnitude;

  const mastersAndLayouts = [].concat(
    presentation.masters || [],
    presentation.layouts || []
  );

  // In dry-run mode, delegate to buildLogoReplaceRequests for logging only.
  if (isDryRun) {
    buildLogoReplaceRequests(mastersAndLayouts, pageWidth, pageHeight, null, true);
    return;
  }

  // Identify objectIds of logo elements using the existing position heuristics.
  // Pass a placeholder URL — we only need the objectIds, not the request objects.
  const requests = buildLogoReplaceRequests(
    mastersAndLayouts, pageWidth, pageHeight, "_placeholder_", false
  );
  if (requests.length === 0) return;

  const logoObjectIds = new Set(
    requests.map(function(r) { return r.replaceImage.imageObjectId; })
  );

  // Fetch the logo blob via DriveApp — no public URL required.
  // SlidesApp.Image.replace(blobSource, crop=false) scales to fit the existing
  // element bounds while preserving aspect ratio (equivalent to CENTER_INSIDE).
  const logoBlob = DriveApp.getFileById(LOGO_CONFIG.newLogoFileId).getBlob();
  const deck = SlidesApp.openById(presentationId);

  deck.getMasters().forEach(function(master) {
    master.getImages().forEach(function(img) {
      if (logoObjectIds.has(img.getObjectId())) img.replace(logoBlob, false);
    });
  });

  deck.getLayouts().forEach(function(layout) {
    layout.getImages().forEach(function(img) {
      if (logoObjectIds.has(img.getObjectId())) img.replace(logoBlob, false);
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
 *   3. Replace Poppins / Figtree fonts with Lexend
 *   4. Replace logo images on master/layout slides
 *
 * @param {string}  presentationId
 * @param {boolean} [dryRun=false]  Passed through to replaceLogos.
 */
function updateSlidesPresentation(presentationId, dryRun) {
  const presentation = Slides.Presentations.get(presentationId);

  Logger.log("Starting brand update for presentation: %s", presentationId);

  // 1. Theme color scheme (masters only)
  updateMasterThemeColors(presentationId, presentation.masters || []);
  Logger.log("  ✓ Master theme colors updated");

  // 2. Inline (direct) colors across all pages
  replaceInlineColors(presentationId);
  Logger.log("  ✓ Inline colors replaced");

  // 3. Fonts across all pages
  replaceFonts(presentationId);
  Logger.log("  ✓ Fonts replaced");

  // 4. Logos on master/layout slides
  replaceLogos(presentationId, dryRun);
  Logger.log("  ✓ Logo replacement %s", dryRun ? "dry run complete" : "complete");
}
