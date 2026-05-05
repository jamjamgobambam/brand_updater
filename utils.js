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
 *
 * newLogoFileId: Google Drive file ID of the replacement logo (used by
 *   SlidesApp.Image.replace via DriveApp blob — file must be shared as
 *   "Anyone with the link can view").
 *
 * newLogoUrl: direct public image URL for the replacement logo. Used by:
 *   - the Docs updater (Docs API cannot follow Drive redirects)
 *   - the Slides delete-and-recreate fallback (createImage requires a URL)
 *
 * slidesLogo: layered detection config for Google Slides.
 *   - oldContentUrlSubstrings: array of substrings to match against an
 *     image element's contentUrl/sourceUrl. Populate by running
 *     logAllImages() against a representative deck and copying a stable
 *     portion of the existing logo's URL. Empty array = URL match disabled.
 *   - zones: named regions (xMin/xMax/yMin/yMax as fractions of slide
 *     dims, 0.0–1.0). An image whose center falls inside a zone AND
 *     whose size/aspect satisfies sizeBounds is treated as a logo.
 *     Empty array = zone fallback disabled (URL-only mode).
 *   - sizeBounds: filter applied ONLY to zone-fallback matches to exclude
 *     hero photos / decorative imagery that happen to sit in a logo zone.
 *     URL matches bypass this filter entirely.
 *
 * docsLogo: legacy schema retained for the Docs updater.
 */
const LOGO_CONFIG = {
<<<<<<< HEAD
  newLogoFileId: "1pIoxLkryTKZjwuWliRp7DQCGKb37F_tU",
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

=======
  newLogoFileId: "1k9CbaVCdgAb5oAfbO5myAG2xH049jGlu",
  cornerLogo: { xThreshold: 0.75, yThreshold: 0.75 },
  titleLogo:  { xMin: 0.25, xMax: 0.75, yMax: 0.35 },
>>>>>>> 11d1b65c6786cbf8973a846826e00292707ea3b2
  docsLogo: {
    oldSourceUri: null, // Set after running logDocImages — e.g. "https://lh3.googleusercontent.com/..."
    // newLogoUrl: direct public image URL for insertInlineImage.
    // The Docs API cannot follow Drive redirects, so drive.google.com URLs fail.
    // Set this to a direct public URL: a GitHub raw URL, Google Cloud Storage,
    // or any CDN that serves the image bytes without redirects.
    // Example: "https://raw.githubusercontent.com/org/repo/main/logo.png"
    // Leave null to fall back to the top-level LOGO_CONFIG.newLogoUrl.
    newLogoUrl:   "https://raw.githubusercontent.com/jamjamgobambam/brand_updater/615367949880121699655c766cb27c68d6206ebe/assets/logo.png",
    minWidthPt:   20,   // Size bounds fallback — adjust based on logDocImages output
    maxWidthPt:   200,
    minHeightPt:  10,
    maxHeightPt:  100,
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
