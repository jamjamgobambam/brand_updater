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
  { oldHex: "#00ADBC", newHex: "#003547" }, // v1 style guide primary teal
  { oldHex: "#00acbc", newHex: "#003547" }, // v1 drawn variant
  { oldHex: "#009eb0", newHex: "#003547" }, // Google Slides theme-slot variant
  { oldHex: "#009eaf", newHex: "#003547" }, // 1-unit variant
  { oldHex: "#0093A4", newHex: "#003547" }, // v2 style guide teal
  { oldHex: "#0093a3", newHex: "#003547" }, // v2 drawn variant

  // ── Accent 2 — Purple ────────────────────────────────────────────────────
  { oldHex: "#7665A0", newHex: "#005E54" }, // v1 style guide purple
  { oldHex: "#7564a0", newHex: "#005E54" }, // v1 drawn variant
  { oldHex: "#9660bf", newHex: "#005E54" }, // Google Slides theme-slot variant
  { oldHex: "#9560bf", newHex: "#005E54" }, // 1-unit variant
  { oldHex: "#8C52BA", newHex: "#005E54" }, // v2 style guide purple
  { oldHex: "#8c52ba", newHex: "#005E54" }, // v2 drawn variant

  // ── Accent 3 ─────────────────────────────────────────────────────────────
  { oldHex: "#ed6060", newHex: "#C2BB00" }, // v1 + v2 style guide strawberry
  { oldHex: "#ED6060", newHex: "#C2BB00" }, // uppercase variant

  // ── Accent 4 ─────────────────────────────────────────────────────────────
  { oldHex: "#3ea33e", newHex: "#E1523D" },
  { oldHex: "#3ea23e", newHex: "#E1523D" }, // 1-unit drawn variant

  // ── Accent 5 — Blue ──────────────────────────────────────────────────────
  { oldHex: "#007acc", newHex: "#ED8B16" }, // Google Slides theme-slot variant
  { oldHex: "#0094CA", newHex: "#ED8B16" }, // v1 style guide blue accent
  { oldHex: "#0093ca", newHex: "#ED8B16" }, // v1 drawn variant

  // ── Accent 6 — Yellow (target is same — normalises near-variants) ─────────
  { oldHex: "#ead300", newHex: "#ead300" },
  { oldHex: "#FFC52D", newHex: "#ead300" }, // v1 style guide bright yellow
  { oldHex: "#ffc42d", newHex: "#ead300" }, // 1-unit drawn variant
];

// Target hex for theme HYPERLINK and FOLLOWED_HYPERLINK slots (same as Accent 2)
const HYPERLINK_NEW_HEX = "#005E54";

/**
 * New hex values for the 6 Accent slots in the master theme ColorScheme,
 * in ACCENT1→ACCENT6 order. Used exclusively by updateMasterThemeColors().
 * Kept separate from COLOR_MAP so adding variant entries to COLOR_MAP never
 * breaks the positional slot assignment.
 */
const ACCENT_NEW_HEXES = [
  "#003547", // ACCENT1 — new teal
  "#005E54", // ACCENT2 — new green
  "#C2BB00", // ACCENT3 — new yellow-green
  "#E1523D", // ACCENT4 — new coral
  "#ED8B16", // ACCENT5 — new orange
  "#ead300", // ACCENT6 — yellow (unchanged)
];

/**
 * Font mapping: old font family → new font family.
 */
const FONT_MAP = [
  { oldFont: "Poppins", newFont: "Lexend" },
  { oldFont: "Figtree", newFont: "Lexend" },
];

/**
 * Font families that are always preserved (never replaced by the fallback).
 * Poppins and Figtree are handled separately via FONT_MAP (→ Lexend) and
 * are intentionally excluded here so the fallback path is never needed for them.
 * Any explicit font NOT in this list and NOT in FONT_MAP will be replaced with FALLBACK_FONT.
 */
const BRAND_FONTS = ["Short Stack", "Lexend"];

/** Replacement font for any non-brand explicit font. */
const FALLBACK_FONT = "Lexend";

/**
 * Euclidean RGB distance threshold (0–255 scale) for near-color matching.
 * Kept tight (15) now that COLOR_MAP has explicit entries for all known
 * old-brand variants. This only covers minor floating-point rounding noise,
 * not fuzzy "similar color" matching.
 */
const COLOR_DISTANCE_THRESHOLD = 15;

/**
 * Logo detection config.
 * newLogoFileId: Google Drive file ID of the replacement logo.
 *   The file must be shared as "Anyone with the link can view".
 * cornerLogo: bottom-right recurring logo (centerX > xThreshold, centerY > yThreshold)
 * titleLogo:  upper-center title slide logo (xMin < centerX < xMax, centerY < yMax)
 * All threshold values are percentages of the slide dimensions (0.0–1.0).
 */
const LOGO_CONFIG = {
  newLogoFileId: "1pIoxLkryTKZjwuWliRp7DQCGKb37F_tU",
  cornerLogo: { xThreshold: 0.75, yThreshold: 0.75 },
  titleLogo:  { xMin: 0.25, xMax: 0.75, yMax: 0.35 },
  docsLogo: {
    oldSourceUri: null, // Set after running logDocImages — e.g. "https://lh3.googleusercontent.com/..."
    // newLogoUrl: direct public image URL for insertInlineImage.
    // The Docs API cannot follow Drive redirects, so drive.google.com URLs fail.
    // Set this to a direct public URL: a GitHub raw URL, Google Cloud Storage,
    // or any CDN that serves the image bytes without redirects.
    // Example: "https://raw.githubusercontent.com/org/repo/main/logo.png"
    // Leave null to fall back to the Drive uc?id= URL (works only if the file
    // is shared "Anyone with the link can view" and Drive serves it redirect-free).
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
