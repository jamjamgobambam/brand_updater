// =============================================================================
// utils.js — Shared constants and helpers for Slides and Docs updaters
// =============================================================================

/**
 * Color mapping: old hex → new hex for each Accent slot.
 * HYPERLINK and FOLLOWED_HYPERLINK both map to the same new hex as Accent 2.
 */
const COLOR_MAP = [
  { oldHex: "#009eb0", newHex: "#003547" }, // Accent 1
  { oldHex: "#9660bf", newHex: "#005E54" }, // Accent 2
  { oldHex: "#ed6060", newHex: "#C2BB00" }, // Accent 3
  { oldHex: "#3ea33e", newHex: "#E1523D" }, // Accent 4
  { oldHex: "#007acc", newHex: "#ED8B16" }, // Accent 5
  { oldHex: "#ead300", newHex: "#ead300" }, // Accent 6 (unchanged)
];

// Target hex for theme HYPERLINK and FOLLOWED_HYPERLINK slots (same as Accent 2)
const HYPERLINK_NEW_HEX = "#005E54";

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
 * Colors within this distance of an old or new brand color will be replaced.
 * 40 ≈ 16% per channel — covers shade variations that are visually the same
 * brand color but stored with slightly different values.
 */
const COLOR_DISTANCE_THRESHOLD = 40;

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
