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
 * Logo detection config.
 * newLogoFileId: Google Drive file ID of the replacement logo.
 *   The file must be shared as "Anyone with the link can view".
 * cornerLogo: bottom-right recurring logo (centerX > xThreshold, centerY > yThreshold)
 * titleLogo:  upper-center title slide logo (xMin < centerX < xMax, centerY < yMax)
 * All threshold values are percentages of the slide dimensions (0.0–1.0).
 */
const LOGO_CONFIG = {
  newLogoFileId: "YOUR_DRIVE_FILE_ID",
  cornerLogo: { xThreshold: 0.75, yThreshold: 0.75 },
  titleLogo:  { xMin: 0.25, xMax: 0.75, yMax: 0.35 },
  docsLogo: {
    oldSourceUri: null, // Set after running logDocImages — e.g. "https://lh3.googleusercontent.com/..."
    minWidthPt:   40,   // Size bounds fallback — adjust based on logDocImages output
    maxWidthPt:   200,
    minHeightPt:  20,
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
 * Builds a publicly accessible download URL for a Google Drive file.
 * The file must be shared as "Anyone with the link can view".
 * @param {string} fileId  Google Drive file ID.
 * @returns {string}
 */
function driveFileUrl(fileId) {
  return `https://drive.google.com/uc?export=download&id=${fileId}`;
}
