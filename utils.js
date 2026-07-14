// =============================================================================
// utils.js — Shared constants and helpers for Slides and Docs updaters
// =============================================================================

/**
 * Color mapping: old hex → new hex for INLINE (direct) RGB color replacement.
 *
 * Each family lists all known stored variants of an old brand color so that
 * slides and docs using either the official style-guide value OR the slightly
 * different value Google rounds/stores in its back-end will both be matched.
 *
 * Sources:
 *   Style guide text values  — extracted from styles1.pdf (v1 brand)
 *   Drawn/rendered variants  — extracted from PDF vector drawings
 *   Style guide text values  — extracted from styles2.pdf (v2 brand)
 *   Google Slides variants   — values Google stores in the theme color slot
 */
const COLOR_MAP = [
  // ── Accent 1 — Teal ──────────────────────────────────────────────────────
  { oldHex: "#00ADBC", newHex: "#6A62D9" }, // v1 style guide primary teal
  { oldHex: "#00acbc", newHex: "#6A62D9" }, // v1 drawn variant
  { oldHex: "#009eb0", newHex: "#6A62D9" }, // Google Slides theme-slot variant
  { oldHex: "#009eaf", newHex: "#6A62D9" }, // 1-unit variant
  { oldHex: "#0093A4", newHex: "#6A62D9" }, // v2 style guide teal
  { oldHex: "#0093a3", newHex: "#6A62D9" }, // v2 drawn variant

  // ── Accent 2 — Purple ────────────────────────────────────────────────────
  { oldHex: "#7665A0", newHex: "#6A62D9" }, // v1 style guide purple
  { oldHex: "#7564a0", newHex: "#6A62D9" }, // v1 drawn variant
  { oldHex: "#9660bf", newHex: "#6A62D9" }, // Google Slides theme-slot variant
  { oldHex: "#9560bf", newHex: "#6A62D9" }, // 1-unit variant
  { oldHex: "#8C52BA", newHex: "#6A62D9" }, // v2 style guide purple
  { oldHex: "#8c52ba", newHex: "#6A62D9" }, // v2 drawn variant

  // ── Accent 3 ─────────────────────────────────────────────────────────────
  { oldHex: "#ed6060", newHex: "#C2BB00" }, // v1 + v2 style guide strawberry
  { oldHex: "#ED6060", newHex: "#C2BB00" }, // uppercase variant

  // ── Accent 4 ─────────────────────────────────────────────────────────────
  { oldHex: "#3ea33e", newHex: "#E1523D" },
  { oldHex: "#3ea23e", newHex: "#E1523D" }, // 1-unit drawn variant

  // ── Accent 5 — Blue ──────────────────────────────────────────────────────
  { oldHex: "#007acc", newHex: "#ED8B16" }, // Google Slides theme-slot variant

  // ── Accent 6 — Yellow (target is same — normalises near-variants) ─────────
  { oldHex: "#ead300", newHex: "#ead300" },
  { oldHex: "#FFC52D", newHex: "#ead300" }, // v1 style guide bright yellow
  { oldHex: "#ffc42d", newHex: "#ead300" }, // 1-unit drawn variant

  // ── Older scheme — all map to new Accent 1 purple ─────────────────────────
  { oldHex: "#0094ca", newHex: "#6A62D9" }, // older Accent 1 (blue)
  { oldHex: "#0094CA", newHex: "#6A62D9" }, // uppercase variant
  { oldHex: "#0093ca", newHex: "#6A62D9" }, // 1-unit drawn variant
  { oldHex: "#ffa400", newHex: "#6A62D9" }, // older Accent 4 (orange)
  { oldHex: "#b9bf15", newHex: "#6A62D9" }, // older Accent 5 (yellow-green)
  { oldHex: "#ffb81d", newHex: "#6A62D9" }, // older Accent 6 (yellow)
];

// Target hex for theme HYPERLINK and FOLLOWED_HYPERLINK slots (same as Accent 2)
const HYPERLINK_NEW_HEX = "#6A62D9";

/**
 * New hex values for the 6 Accent slots in the master theme ColorScheme,
 * in ACCENT1→ACCENT6 order. Used exclusively by updateMasterThemeColors().
 * Kept separate from COLOR_MAP so adding variant entries to COLOR_MAP never
 * breaks the positional slot assignment.
 */
const ACCENT_NEW_HEXES = [
  "#6A62D9", // ACCENT1 — new purple
  "#6A62D9", // ACCENT2 — new purple (same as Accent 1)
  "#C2BB00", // ACCENT3 — new yellow-green
  "#E1523D", // ACCENT4 — new coral
  "#ED8B16", // ACCENT5 — new orange
  "#ead300", // ACCENT6 — yellow (unchanged)
];

/**
 * Font mapping: old font family → new font family.
 */
const FONT_MAP = [
  { oldFont: "Poppins", newFont: "Geist" },
  { oldFont: "Figtree", newFont: "Geist" },
];

/**
 * Font families that are always preserved (never replaced by the fallback).
 * Poppins and Figtree are handled separately via FONT_MAP (→ Geist) and
 * are intentionally excluded here so the fallback path is never needed for them.
 * Any explicit font NOT in this list and NOT in FONT_MAP will be replaced with FALLBACK_FONT.
 */
const BRAND_FONTS = ["Short Stack", "Geist"];

/** Replacement font for any non-brand explicit font. */
const FALLBACK_FONT = "Geist";

/**
 * Euclidean RGB distance threshold (0–255 scale) for near-color matching.
 * Kept tight (15) now that COLOR_MAP has explicit entries for all known
 * old-brand variants. This only covers minor floating-point rounding noise,
 * not fuzzy "similar color" matching.
 */
const COLOR_DISTANCE_THRESHOLD = 15;

/**
 * Logo detection config.
 * newLogoFileId: Google Drive file ID of the replacement logo. Used as a
 *   fallback only — the preferred source is the CODEAI_LOGO_DRIVE_ID Script
 *   Property (read via getCodeAILogoFileId_). The file must be shared as
 *   "Anyone with the link can view".
 * cornerLogo: bottom-right recurring logo (centerX > xThreshold, centerY > yThreshold)
 * titleLogo:  upper-center title slide logo (xMin < centerX < xMax, centerY < yMax)
 * All threshold values are percentages of the slide dimensions (0.0–1.0).
 */
const LOGO_CONFIG = {
  newLogoUrl:    "https://raw.githubusercontent.com/jamjamgobambam/brand_updater/615367949880121699655c766cb27c68d6206ebe/assets/logo.png",

  slidesLogo: {
    // Populate after running logAllImages() — e.g. ["lh3.googleusercontent.com/abc123"]
    // or a stable portion of the original sourceUrl. Empty = skip URL matching.
    oldContentUrlSubstrings: [],

    // Named regions on the slide. centerX/centerY of the image must fall
    // inside (xMin..xMax, yMin..yMax). Values are fractions of slide dims.
    //
    // DEFAULT: empty array = URL-only mode. Position fallback is OFF by
    // default to avoid replacing unrelated imagery (e.g. toggle/button
    // illustrations in lesson decks) on presentations where the original
    // logo's URL has not yet been identified. Populate this array only
    // after logAllImages() confirms the safe zones for a given template.
    zones: [],

    // Reference zone defaults — copy individual entries into `zones`
    // above when enabling position fallback for a known template.
    zonesReference: [
      { name: "bottom-right",  xMin: 0.75, xMax: 1.00, yMin: 0.75, yMax: 1.00 },
      { name: "bottom-left",   xMin: 0.00, xMax: 0.25, yMin: 0.75, yMax: 1.00 },
      { name: "bottom-center", xMin: 0.25, xMax: 0.75, yMin: 0.75, yMax: 1.00 },
      { name: "top-left",      xMin: 0.00, xMax: 0.25, yMin: 0.00, yMax: 0.35 },
      { name: "top-right",     xMin: 0.75, xMax: 1.00, yMin: 0.00, yMax: 0.35 },
      { name: "top-center",    xMin: 0.25, xMax: 0.75, yMin: 0.00, yMax: 0.35 },
    ],

    // Size/aspect filter for zone-fallback matches only. Width/height as
    // fractions of slide dims; aspect = width / height.
    sizeBounds: {
      minWidthPct:  0.02,
      maxWidthPct:  0.40,
      minHeightPct: 0.02,
      maxHeightPct: 0.40,
      minAspect:    0.20,
      maxAspect:    8.00,
    },
  },

  newLogoFileId: "1k9CbaVCdgAb5oAfbO5myAG2xH049jGlu",
  // cornerLogo / titleLogo position thresholds are no longer used for slides
  // (Gemini-based classifier handles every image), but kept for reference and
  // potential reuse by the docs path.
  cornerLogo: { xThreshold: 0.75, yThreshold: 0.75 },
  titleLogo:  { xMin: 0.25, xMax: 0.75, yMax: 0.35 },
  // Top-left header band — legacy decks render a small wordmark in the
  // top-left corner of every layout, with the course name immediately to
  // its right. After replacement, the new (roughly square) logo overlaps
  // that adjacent text unless we shift the text right. Used by
  // replaceSlidesLogoImage_ to detect this case.
  headerLogo: { xMax: 0.15, yMax: 0.15 },
  docsLogo: {
    oldSourceUri: null, // Set after running logDocImages — e.g. "https://lh3.googleusercontent.com/..."
    // Additional known sourceUris of legacy logo images. Any inline image
    // whose embeddedObject.imageProperties.sourceUri is in this list (or
    // matches oldSourceUri) is treated as a logo and replaced. Cheaper
    // than the Gemini classifier and useful for fast paths on docs whose
    // logo URI is known.
    oldSourceUris: [],
    // When true and no sourceUri match is found, fetch the image bytes
    // (via embeddedObject.imageProperties.contentUri + OAuth) and run the
    // Gemini logo classifier (logo-classifier.js → classifyLogo_). Only
    // images classified as "replace" become matches. This avoids
    // false-positive replacements of unrelated images (avatars, diagrams,
    // screenshots) that happen to share size bounds with the legacy logo.
    useGeminiClassifier: true,
    // newLogoUrl: optional override — direct public image URL for insertInlineImage.
    // The Docs API cannot follow Drive redirects, so drive.google.com URLs
    // sometimes fail; if you have a CDN/raw URL, set it here to override.
    // Leave null (default) to fall back to the Drive uc?id= URL built from
    // the CODEAI_LOGO_DRIVE_ID Script Property (or LOGO_CONFIG.newLogoFileId).
    newLogoUrl:   null,
    minWidthPt:   20,   // Size bounds — used ONLY by logDocLogoCandidates as a
    maxWidthPt:   200,  // discovery filter, not by the live matching path.
    minHeightPt:  10,
    maxHeightPt:  100,
    // Uniform scale factor applied to the replacement logo's size. The new
    // CODEAI logo is much shorter than the legacy logo, so reusing the old
    // box dimensions makes it look tiny. Scaling the inserted image up (4x
    // by default ≈ 400% wider) restores visual prominence while preserving
    // the new logo's natural aspect ratio inside the larger box.
    scale:        4.0,
    // Resize columns of any table that contains a matched logo, in inches.
    // Index 0 → first column width, index 1 → second column width, etc.
    // Columns beyond the array's length are left unchanged. The legacy
    // template uses a 1.1" / 6.2" split that crowds the logo against the
    // right edge; flipping to 6.2" / 2.3" gives the scaled logo room.
    tableColumnWidthsIn: [6.2, 2.3],
    // When true, replaceDocLogos runs a structural pass that deletes a
    // legacy empty leading spacer column from 3-column logo tables and
    // applies tableColumnWidthsIn. Set to false to disable structural
    // changes (only the logo image itself is replaced) — useful for
    // diagnosing layout issues or running on docs whose tables shouldn't
    // be restructured.
    restructureLegacyTable: true,
    // Hard maximum width for the inserted (replacement) logo, in inches.
    // The image is also clamped to its containing cell / page content
    // width; this cap applies on top so the logo never exceeds the brand
    // size regardless of cell width or scale factor. Set to null to
    // disable the global cap.
    maxWidthInsertedIn: 2.0,
    // Left indent (inches) applied to every paragraph in column 0 of any
    // table touched by the structural pass. Keeps title text from
    // touching the cell's left edge after the spacer column is removed.
    // Set to null or 0 to disable.
    firstColumnTextIndentIn: 0.5,
  },
};

// =============================================================================
// Helper functions
// =============================================================================

/**
 * Converts a "#RRGGBB" hex string to a normalized RGB object { red, green, blue }
 * with component values in the range 0.0–1.0, as required by the Slides REST API.
 * @param {string} hex  Six-digit hex color string, with or without leading "#".
 * @returns {{ red: number, green: number, blue: number }}
 */
function hexToNormalizedRgb(hex) {
  const clean = hex.replace(/^#/, "");
  return {
    red:   parseInt(clean.substring(0, 2), 16) / 255,
    green: parseInt(clean.substring(2, 4), 16) / 255,
    blue:  parseInt(clean.substring(4, 6), 16) / 255,
  };
}

/**
 * Compares an API rgbColor object against a target hex string.
 * Uses a float tolerance of 1/255 (~0.004) to account for rounding in
 * Google's backend storage.
 * @param {{ red?: number, green?: number, blue?: number }} apiRgb
 * @param {string} targetHex  Six-digit hex color string.
 * @param {number} [tolerance=1/255]
 * @returns {boolean}
 */
function normalizedRgbMatches(apiRgb, targetHex, tolerance) {
  if (!apiRgb) return false;
  const tol = tolerance !== undefined ? tolerance : 1 / 255;
  const target = hexToNormalizedRgb(targetHex);
  return (
    Math.abs((apiRgb.red   || 0) - target.red)   <= tol &&
    Math.abs((apiRgb.green || 0) - target.green) <= tol &&
    Math.abs((apiRgb.blue  || 0) - target.blue)  <= tol
  );
}

/**
 * Computes the Euclidean distance between an API rgbColor object and a target
 * hex color in 0–255 RGB space.
 * @param {{ red?: number, green?: number, blue?: number }} apiRgb
 * @param {string} targetHex  Six-digit hex color string.
 * @returns {number}  Distance in 0–255 space; Infinity if apiRgb is falsy.
 */
function colorDistance(apiRgb, targetHex) {
  if (!apiRgb) return Infinity;
  var target = hexToNormalizedRgb(targetHex);
  var dr = ((apiRgb.red   || 0) - target.red)   * 255;
  var dg = ((apiRgb.green || 0) - target.green) * 255;
  var db = ((apiRgb.blue  || 0) - target.blue)  * 255;
  return Math.sqrt(dr * dr + dg * dg + db * db);
}

/**
 * Finds the replacement hex for an API rgbColor by range-matching against
 * colorMap entries using Euclidean distance in 0–255 RGB space.
 *
 * Matching priority:
 *   1. Within `threshold` of an OLD brand color → returns that entry's newHex.
 *   2. Within `threshold` of a NEW brand color  → snaps to that exact newHex.
 *
 * @param {{ red?: number, green?: number, blue?: number }} apiRgb
 * @param {Object[]} colorMap   Array of { oldHex, newHex } pairs.
 * @param {number}   [threshold=COLOR_DISTANCE_THRESHOLD]
 * @returns {string|null}  newHex to use, or null if no match.
 */
function findColorMapping(apiRgb, colorMap, threshold) {
  if (!apiRgb) return null;
  var thr = threshold !== undefined ? threshold : COLOR_DISTANCE_THRESHOLD;
  // Pass 1: near an old brand color
  for (var i = 0; i < colorMap.length; i++) {
    if (colorDistance(apiRgb, colorMap[i].oldHex) <= thr) {
      return colorMap[i].newHex;
    }
  }
  // Pass 2: near a new brand color — snap to exact new value
  for (var j = 0; j < colorMap.length; j++) {
    if (colorDistance(apiRgb, colorMap[j].newHex) <= thr) {
      return colorMap[j].newHex;
    }
  }
  return null;
}

/**
 * Fetches a presentation with up to maxAttempts retries on transient errors
 * (e.g. "Empty response"). Waits 2^attempt seconds between retries.
 *
 * @param {string} presentationId
 * @param {number} [maxAttempts=3]
 * @returns {Object} Presentation resource from the Slides API.
 */
function getPresentation(presentationId, maxAttempts) {
  const attempts = maxAttempts || 3;
  for (var i = 0; i < attempts; i++) {
    try {
      return Slides.Presentations.get(presentationId);
    } catch (e) {
      if (i === attempts - 1) throw e;
      Utilities.sleep(Math.pow(2, i) * 1000);
    }
  }
}

// =============================================================================
// Gemini-based logo classifier — config & shared helpers
// =============================================================================

/**
 * Script Property keys read at runtime. Set via:
 *   Apps Script editor > Project Settings > Script properties.
 *
 * GEMINI_API_KEY is required for slides logo replacement; without it,
 * the classifier short-circuits to "skip" and no logos are replaced.
 */
const PROP_KEYS = {
  GEMINI_API_KEY: "GEMINI_API_KEY",
  CODEAI_LOGO_DRIVE_ID: "CODEAI_LOGO_DRIVE_ID",
};

/**
 * Confidence thresholds for the Gemini logo classifier:
 *   confidence >= REPLACE          → replace the image with the new logo
 *   REVIEW <= confidence < REPLACE → leave alone, surface as needs-review
 *   confidence < REVIEW            → ignore
 */
const LOGO_CONFIDENCE = {
  REPLACE: 0.85,
  REVIEW:  0.5,
};

/**
 * Default uniform scale applied to a replaced logo's WIDTH (height follows
 * the new logo's natural aspect ratio). Used as a fallback when the original
 * doesn't match the title-center or bottom-right corner position bands.
 */
const LOGO_SCALE = 2;

/**
 * Width-scale applied to logos detected in the title-slide top-center band
 * (centerX in [0.25, 0.75], centerY < 0.45). The new logo is roughly square
 * while the legacy wordmark was wide-and-short, so scaling by the original's
 * (small) height made the replacement look tiny. Scaling the WIDTH directly
 * — and letting the height follow the new logo's natural aspect — restores
 * visual prominence on title slides.
 */
const TITLE_LOGO_WIDTH_SCALE = 3;

/**
 * Minimum WIDTH (in points; 72pt = 1 inch) for logos detected in the
 * bottom-right corner band (centerX > 0.75, centerY > 0.75). If the original
 * is wider than this, the original width is preserved.
 */
const CORNER_LOGO_MIN_WIDTH_PT = 72;

/**
 * Minimum WIDTH (in points; 36pt = 0.5 inch) for logos detected in the
 * top-left header band (LOGO_CONFIG.headerLogo). If the original is wider
 * than this, the original width is preserved — header logos are usually
 * sized correctly already; the issue is overlap with adjacent text, not
 * size.
 */
const HEADER_LOGO_MIN_WIDTH_PT = 36;

/**
 * Horizontal gap (in points) inserted between the new top-left header logo's
 * right edge and any text shape that gets shifted out of its way.
 */
const HEADER_TEXT_GAP_PT = 8;

/**
 * Minimum padding (in points; Slides' native unit) between the new logo and
 * any slide edge. If centering would put the new logo closer to an edge than
 * this, it is shifted inward.
 */
const LOGO_SLIDE_MARGIN = 10;

/** Gemini model used for the vision-based logo classifier. */
const GEMINI_MODEL = "gemini-2.5-flash";

/**
 * Reads a Script Property; returns null when the key is unset so callers can
 * decide how to fail (the classifier short-circuits to "skip").
 */
function getScriptProp_(key) {
  try {
    return PropertiesService.getScriptProperties().getProperty(key);
  } catch (e) {
    return null;
  }
}

function getGeminiKey_() {
  return getScriptProp_(PROP_KEYS.GEMINI_API_KEY);
}

/**
 * Returns the Drive file ID of the replacement CodeAI logo.
 * Prefers the CODEAI_LOGO_DRIVE_ID Script Property; falls back to the
 * hard-coded LOGO_CONFIG.newLogoFileId so existing deployments keep working.
 */
function getCodeAILogoFileId_() {
  return getScriptProp_(PROP_KEYS.CODEAI_LOGO_DRIVE_ID) || LOGO_CONFIG.newLogoFileId;
}

/** Cached blob for the new CodeAI logo so DriveApp is hit at most once per run. */
var _codeAiLogoBlobCache = null;
function getCodeAILogoBlob_() {
  if (_codeAiLogoBlobCache) return _codeAiLogoBlobCache;
  _codeAiLogoBlobCache = DriveApp.getFileById(getCodeAILogoFileId_()).getBlob();
  return _codeAiLogoBlobCache;
}

/** Lowercase hex SHA-256 of a blob's bytes (used as the classifier cache key). */
function sha256Hex_(blob) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    blob.getBytes()
  );
  let s = "";
  for (let i = 0; i < bytes.length; i++) {
    const b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
    s += (b < 16 ? "0" : "") + b.toString(16);
  }
  return s;
}
