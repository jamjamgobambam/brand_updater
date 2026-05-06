// =============================================================================
// logo-classifier.js — Gemini-based logo detection for slide images
//
// Classifier results are cached per-image in Document/Script Properties keyed
// by the SHA-256 of the image bytes, so repeated runs and repeated images
// within one file do not re-call Gemini.
//
// Depends on globals defined in utils.js:
//   GEMINI_MODEL, LOGO_CONFIDENCE, getGeminiKey_, sha256Hex_
// =============================================================================

/**
 * Property store used to cache classifier verdicts. DocumentProperties is
 * unavailable from a standalone (non-bound) script; ScriptProperties is the
 * fallback. Both share the same get/setProperty surface.
 */
const _logoCacheStore = (function() {
  try {
    return PropertiesService.getDocumentProperties() || PropertiesService.getScriptProperties();
  } catch (e) {
    return PropertiesService.getScriptProperties();
  }
})();

/**
 * Classify a single image blob. Returns { is_logo: boolean, confidence: number }.
 * Cached by SHA-256 of the image bytes. Bump the cache version (`logo_v…`) when
 * the prompt or thresholds change so prior verdicts don't override new logic.
 *
 * @param {GoogleAppsScript.Base.Blob} blob
 * @returns {{ is_logo: boolean, confidence: number }}
 */
function classifyLogo_(blob) {
  const hash = sha256Hex_(blob);
  const cacheKey = "logo_v2_" + hash;

  const cached = _logoCacheStore.getProperty(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through to refetch */ }
  }

  const result = callGeminiVision_(blob);
  try {
    _logoCacheStore.setProperty(cacheKey, JSON.stringify(result));
  } catch (e) {
    // Properties have a 500KB cap; ignore overflow — classification still works.
  }
  return result;
}

/**
 * Calls the Gemini vision API with a strict logo-classifier prompt.
 * Returns { is_logo: false, confidence: 0 } on any error or missing API key
 * so the caller can safely treat it as "skip".
 *
 * @param {GoogleAppsScript.Base.Blob} blob
 * @returns {{ is_logo: boolean, confidence: number }}
 */
function callGeminiVision_(blob) {
  const apiKey = getGeminiKey_();
  if (!apiKey) {
    Logger.log(
      "Gemini API key missing — set Script Property %s to enable logo replacement.",
      "GEMINI_API_KEY"
    );
    return { is_logo: false, confidence: 0 };
  }

  const mime = blob.getContentType() || "image/png";
  const b64  = Utilities.base64Encode(blob.getBytes());

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(GEMINI_MODEL) +
    ":generateContent?key=" + encodeURIComponent(apiKey);

  const payload = {
    contents: [{
      role: "user",
      parts: [
        { text:
          "You are a strict brand-logo classifier. Return is_logo:true ONLY " +
          'if the image is a STANDALONE code.org logo: the "code.org" ' +
          "wordmark or its geometric mark, presented as a logo asset (any " +
          "color treatment, including monochrome), with NO surrounding UI, " +
          "document content, slide content, or photographic context.\n\n" +
          "Return is_logo:false for: screenshots of code.org pages, slide " +
          "thumbnails, activity-guide screenshots, photos, or any composite " +
          "image that merely contains a code.org logo somewhere within it. " +
          "Containing the logo is not enough; the entire image must BE the " +
          "logo.\n\n" +
          "Respond ONLY with JSON: " +
          '{"is_logo": <bool>, "confidence": <number between 0 and 1>}. ' +
          "Be conservative: only return confidence above 0.9 when the image " +
          "is unambiguously a code.org logo asset."
        },
        { inlineData: { mimeType: mime, data: b64 } },
      ],
    }],
    generationConfig: {
      temperature: 0,
      responseMimeType: "application/json",
    },
  };

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    Logger.log("Gemini vision call failed: HTTP %s — %s", code, resp.getContentText());
    return { is_logo: false, confidence: 0 };
  }

  try {
    const body = JSON.parse(resp.getContentText());
    const text =
      body.candidates && body.candidates[0] &&
      body.candidates[0].content && body.candidates[0].content.parts &&
      body.candidates[0].content.parts[0] && body.candidates[0].content.parts[0].text;
    if (!text) return { is_logo: false, confidence: 0 };
    const parsed = JSON.parse(text);
    return {
      is_logo: !!parsed.is_logo,
      confidence: Math.max(0, Math.min(1, Number(parsed.confidence) || 0)),
    };
  } catch (err) {
    Logger.log("Gemini vision parse failed: %s", err && err.message);
    return { is_logo: false, confidence: 0 };
  }
}

/**
 * Decide what to do with a classifier result.
 * Returns one of: 'replace' | 'review' | 'skip'.
 */
function logoAction_(result) {
  if (!result || !result.is_logo) return "skip";
  if (result.confidence >= LOGO_CONFIDENCE.REPLACE) return "replace";
  if (result.confidence >= LOGO_CONFIDENCE.REVIEW)  return "review";
  return "skip";
}
