# Plan Part 3: Docs Theme, Color, Font and Logo Updater

## Summary

Update Google Docs documents by:

1. Traversing all text runs with explicit (inline) color overrides and replacing old brand colors with new ones
2. Traversing all text runs with explicit font overrides and replacing Poppins and Figtree with Lexend
3. Finding and replacing logo images in both the document body and headers/footers

Uses the **Advanced Docs REST API** (`Docs.Documents.batchUpdate()`). A single-file function handles one document; a batch wrapper handles multiple from a Drive folder.

> **Named Styles limitation:** The Docs REST API has no `updateNamedStyles` request — it is not possible to programmatically change Named Style defaults (e.g., Heading 1, Normal Text). Only explicit inline overrides applied on top of Named Styles can be updated. Named Style defaults in template documents must be updated manually.

---

## Color Mapping

Same as Slides — shared from `utils.js`:

| Accent slot | Old hex | New hex |
|---|---|---|
| Accent 1 | #009eb0 | #003547 |
| Accent 2, Link, and Followed Hyperlink | #9660bf | #005E54 |
| Accent 3 | #ed6060 | #C2BB00 |
| Accent 4 | #3ea33e | #E1523D |
| Accent 5 | #007acc | #ED8B16 |
| Accent 6 | #ead300 | #ead300 (unchanged) |

---

## Font Mapping

Same as Slides — shared from `utils.js`:

| Old font | New font |
|---|---|
| Poppins | Lexend |
| Figtree | Lexend |

---

## Logo Config

Primary matching strategy: `sourceUri` comparison. Unlike Slides (where logos are typically uploaded from Drive and `sourceUrl` is often empty), Docs inline images tend to retain the `sourceUri` from the Drive file. If `sourceUri` is unavailable or null, fall back to size bounds matching.

Logo location: both document **body** and **headers/footers**.

---

## Phase 0 — Refactor Shared Utilities into `utils.js`

> **This phase must be completed before writing `docs-updater.js`.** It extracts constants and helpers shared between Slides and Docs into a single global file.

### Step 1 — Create `utils.js`

Move the following from `slides-updater.js` into a new `utils.js` file:

- `COLOR_MAP` — array of `{ oldHex, newHex }` pairs
- `FONT_MAP` — array of `{ oldFont, newFont }` pairs
- `hexToNormalizedRgb(hex)` — hex → normalized float RGB
- `normalizedRgbMatches(apiRgb, targetHex, tolerance)` — float comparison with tolerance
- `driveFileUrl(fileId)` — builds Drive export URL
- `LOGO_CONFIG` — logo detection thresholds and Drive file IDs (expand to include a `docsLogo` key alongside `cornerLogo` and `titleLogo`)

In Apps Script, all `.js` files share a single global scope at runtime (they are concatenated). `utils.js` defined in the project is automatically available to `slides-updater.js` and `docs-updater.js` without any import statements.

---

### Step 2 — Remove shared items from `slides-updater.js`

Delete the definitions of `COLOR_MAP`, `FONT_MAP`, `hexToNormalizedRgb`, `normalizedRgbMatches`, `driveFileUrl`, and `LOGO_CONFIG` from `slides-updater.js`. Because the global scope is shared, all references in `slides-updater.js` continue to resolve correctly from `utils.js`.

This prevents duplicate declarations. In Apps Script's global scope, defining the same `const` or `function` name in two files causes a runtime error.

---

## Phase 1 — Manifest

### Step 3 — Add Docs Advanced Service

Add to `appsscript.json` under `dependencies.enabledAdvancedServices`:

```json
{ "userSymbol": "Docs", "serviceId": "docs", "version": "v1" }
```

Also enable it in the cloud project via the Apps Script editor → Services panel.

Same two-step requirement as the Slides Advanced API. The basic `DocumentApp` service does not expose the full REST API — you need the Advanced Docs service for `batchUpdate`, which is required for structural changes like replacing text style colors and fonts. If not declared, the `Docs` global object won't exist at runtime.

---

## Phase 2 — Implement `docs-updater.js`

### Step 4 — `collectDocContent(document)`

Returns a flat array of `{ content, segmentId }` pairs, one per segment of the document:

- Body: `{ content: document.body.content, segmentId: "" }`
- Each named header: `{ content: document.headers[headerId].content, segmentId: headerId }` — iterate all keys of `document.headers`
- Each named footer: `{ content: document.footers[footerId].content, segmentId: footerId }` — iterate all keys of `document.footers`
- Each footnote: `{ content: document.footnotes[footnoteId].content, segmentId: footnoteId }` — iterate all keys of `document.footnotes`; skip if the key is absent (documents with no footnotes omit it entirely)

**Why `segmentId`:** Every range-based request in the Docs API requires a `segmentId` field. For the body it is an empty string `""`; headers, footers, and footnotes each have their own opaque ID assigned by Google. Failing to pass the correct `segmentId` causes the API to reject the request with a 400 error.

**Why collect all segments upfront:** Having a uniform `{ content, segmentId }` array means color, font, and logo traversal functions share one loop structure without needing to know about body vs. header vs. footer vs. footnote differences. It also ensures headers, footers, and footnotes (which can contain styled text) are not skipped.

---

### Step 5 — `buildDocColorRequests(segments, colorMap)`

Core traversal for color replacement. For each segment, iterates the `content` array:

- For paragraphs: `element.paragraph.elements[].textRun.textStyle.foregroundColor.color.rgbColor`
- For table cells: `element.table.tableRows[].tableCells[].content[].paragraph.elements[].textRun.textStyle.foregroundColor.color.rgbColor`

For each `rgbColor` found, call `normalizedRgbMatches()` against each entry in `colorMap`. On a match, build an `updateTextStyle` request:

```json
{
  "updateTextStyle": {
    "range": {
      "startIndex": <textRun startIndex>,
      "endIndex": <textRun endIndex>,
      "segmentId": <segment's segmentId>
    },
    "textStyle": {
      "foregroundColor": {
        "color": { "rgbColor": { "red": ..., "green": ..., "blue": ... } }
      }
    },
    "fields": "foregroundColor"
  }
}
```

Returns array of `updateTextStyle` request objects.

**Why `fields: "foregroundColor"` only:** The Docs API uses a field mask to determine which properties to overwrite. Specifying only `"foregroundColor"` preserves all other text style properties (bold, italic, font size, font family, etc.) — the API ignores anything not listed.

**Why `startIndex`/`endIndex` per text run:** A single paragraph can contain several text runs with different styles. Each `textRun` has a `startIndex` and `endIndex` within the segment. You must target the exact run, not the whole paragraph, to avoid overwriting the style of adjacent text with a different color.

**This implements Option A**: only explicit `foregroundColor` overrides are updated. Text that displays a color because of its Named Style (e.g., default Normal Text color) has no `rgbColor` here and is not touched.

---

### Step 6 — `replaceDocColors(docId)`

*Depends on steps 4, 5.*

- Calls `Docs.Documents.get(docId)` — full document JSON in one read
- Calls `collectDocContent(document)` to get all segments
- Calls `buildDocColorRequests(segments, COLOR_MAP)` to build requests
- If requests are non-empty, submits via `Docs.Documents.batchUpdate(docId, { requests })`

One read, one write — same pattern as the Slides functions. The Docs API rate limits batchUpdate calls per document, so batching all color changes into a single request is both more efficient and less likely to hit quota limits.

---

### Step 7 — `buildDocFontRequests(segments, fontMap)`

Same traversal structure as `buildDocColorRequests`. For each `textRun`:

- Read `textStyle.weightedFontFamily.fontFamily` (preferred — includes weight)
- Fall back to `textStyle.fontFamily` if `weightedFontFamily` is absent
- If either matches an old font name in `fontMap`, build an `updateTextStyle` request with:
  - `textStyle.weightedFontFamily: { fontFamily: newFont, weight: <existing weight> }`
  - `fields: "weightedFontFamily"`

Returns array of `updateTextStyle` request objects.

**Why `weightedFontFamily` (not just `fontFamily`):** Same reason as Slides — the API stores font and weight together. Writing only `fontFamily` silently resets weight to normal (400), turning bold Lexend runs into regular weight. Preserving the existing weight prevents this regression.

**Why `fields: "weightedFontFamily"` only:** Other text properties (color, size, italic) are not part of this update and should not be touched.

---

### Step 8 — `replaceDocFonts(docId)`

*Depends on steps 4, 7.*

- Calls `Docs.Documents.get(docId)`
- Calls `collectDocContent(document)` to get all segments
- Calls `buildDocFontRequests(segments, FONT_MAP)` to build requests
- If requests are non-empty, submits via `Docs.Documents.batchUpdate(docId, { requests })`

---

### Step 9 — `logDocImages(docId)`

Diagnostic/discovery utility. Run this **once on a representative document before tuning `LOGO_CONFIG.docsLogo`**.

- Calls `Docs.Documents.get(docId)` to get the full document JSON
- Iterates `document.inlineObjects` (a map of `objectId → inlineObject`)
- For each inline object, reads `inlineObjectProperties.embeddedObject`:
  - `imageProperties.sourceUri` — the original image URL (often populated in Docs)
  - `size.width.magnitude` and `.unit` — dimensions in the document's native unit (usually `PT`)
  - `size.height.magnitude`
- Logs: `objectId`, `sourceUri`, width, height, and which segment it appears in (body, header, footer)
- Does not make any changes

**Why:** Before configuring `LOGO_CONFIG.docsLogo`, you need to know whether `sourceUri` is populated (enabling direct URL matching) and what the dimensions of existing logos are (so you can set appropriate `minWidthPt`/`maxWidthPt`/`minHeightPt`/`maxHeightPt` bounds as a fallback). Running this once gives you that data from real documents.

The `inlineObjects` map is the Docs API's way of storing images. Each entry is referenced from the document body/header/footer content via an `inlineObjectElement` that contains the `inlineObjectId`. This is the structure you'll traverse in later steps.

---

### Step 10 — `LOGO_CONFIG.docsLogo` (in `utils.js`)

Add a `docsLogo` key to the existing `LOGO_CONFIG` object in `utils.js`:

```js
docsLogo: {
  oldSourceUri: null,      // Set after running logDocImages — e.g. "https://lh3.googleusercontent.com/..."
  minWidthPt: 40,          // Size bounds fallback — adjust based on logDocImages output
  maxWidthPt: 200,
  minHeightPt: 20,
  maxHeightPt: 100
}
```

`oldSourceUri` is `null` until you run `logDocImages` and identify the correct URI from the output. Once set, it becomes the primary match criterion. The size bounds serve as a fallback for logos uploaded directly (not from a URL) where `sourceUri` is null.

---

### Step 11 — `buildDocLogoRequests(document, newLogoUrl, dryRun)`

Finds all logo inline objects and builds delete + re-insert request pairs.

Process:

1. Build a reverse index: for each `inlineObjectElement` in all segments (body + headers + footers), record:
   - `objectId`
   - `startIndex` of the `inlineObjectElement`
   - `segmentId` of its containing segment
2. For each entry in the reverse index, look up `document.inlineObjects[objectId]` to get `sourceUri` and dimensions
3. Match criteria (checking in order):
   - If `LOGO_CONFIG.docsLogo.oldSourceUri` is non-null: match if `sourceUri === oldSourceUri`
   - Otherwise (fallback): match if width and height are within `minWidthPt`/`maxWidthPt`/`minHeightPt`/`maxHeightPt`
4. On a match:
   - If `dryRun`: log `objectId`, `segmentId`, `startIndex`, `sourceUri`, dimensions — no request built
   - If not `dryRun`: capture the existing logo's `embeddedObject.size` (width and height in PT), then build a **pair** of requests:
     1. `deleteContentRange` at `{ startIndex, endIndex: startIndex + 1, segmentId }`
     2. `insertInlineImage` at `{ index: startIndex, segmentId }` with `uri: newLogoUrl` and `objectSize: { width: { magnitude: <existing width>, unit: "PT" }, height: { magnitude: <existing height>, unit: "PT" } }`
5. **Sort all pairs in reverse `startIndex` order** before returning

Returns flat array of request objects (delete + insert interleaved), sorted reverse by index.

**Why delete + insert:** The Docs API has no `replaceInlineImage` equivalent. To swap a logo, you delete the `inlineObjectElement` (which removes the image from the document flow) and insert a new inline image at the same position. Unlike Slides' `replaceImage` (which preserves the bounding box automatically), `insertInlineImage` inserts at the image's natural size unless you pass `objectSize` — so the existing dimensions must be captured before deletion and passed explicitly to preserve the logo's layout.

**Why reverse index order:** When you delete content at a `startIndex` and insert at the same position, the indices of all subsequent elements shift. Processing in reverse order (highest index first) means earlier operations don't invalidate the indices of later operations. If done in ascending order, the first deletion would shift all subsequent indices by -1 (or +1 after insertion), causing every subsequent operation to target the wrong location.

**Why build a reverse index first:** The Docs API represents document structure as `content[]` arrays within each segment. To find the `startIndex` of an `inlineObjectElement`, you must traverse the content. The `document.inlineObjects` map gives you image properties but not positions — you must find positions by traversing content and correlating `inlineObjectId` values.

---

### Step 12 — `replaceDocLogos(docId, dryRun)`

*Depends on steps 10, 11. (Step 9 is a one-time configuration prerequisite — run it to calibrate the `LOGO_CONFIG.docsLogo` values used by step 10, but it is not called by this function.)* 

- Calls `Docs.Documents.get(docId)` to get the full document JSON
- Calls `driveFileUrl(LOGO_CONFIG.newLogoFileId)` to build the new logo URL
- Calls `buildDocLogoRequests(document, newLogoUrl, dryRun)`
- If not dry run and requests are non-empty, submits via `Docs.Documents.batchUpdate(docId, { requests })`

---

### Step 13 — `updateDocsDocument(docId, dryRun)`

Public orchestrator. Sequences:

1. `replaceDocColors(docId)`
2. `replaceDocFonts(docId)`
3. `replaceDocLogos(docId, dryRun)`

Kept separate from steps 6, 8, and 12 for testability and reuse — each function can be called and tested independently. Colors run first, then fonts, then logos. The `dryRun` flag is passed through to `replaceDocLogos`.

---

## Phase 3 — Batch Wrapper in `main.js`

### Step 14 — `updateAllDocsInFolder(folderId, dryRun)`

Entry point for batch runs:

- Uses `DriveApp.getFolderById(folderId).getFiles()` to iterate files
- Filters by MIME type `application/vnd.google-apps.document`
- Wraps each `updateDocsDocument(file.getId(), dryRun)` call in a `try/catch` — logs the file name and error message on failure and continues to the next file rather than aborting the entire batch
- Logs progress (file name, success or failure) with `Logger.log()`

`application/vnd.google-apps.document` is Google's MIME type for Docs files. Same Drive iteration pattern as `updateAllSlidesInFolder` in `main.js` — consistent interface for both updaters.

**Apps Script execution timeout:** Apps Script hard-kills execution after 6 minutes (30 minutes for Google Workspace accounts). A large folder could exceed this limit. If that happens, run the updater on smaller subfolders. A more robust future improvement would be to store a file-index cursor in `PropertiesService` so a subsequent run can resume from where the previous one stopped.

---

## Files to Create / Modify

| File | Action | Purpose |
|---|---|---|
| `utils.js` | **Create** | Shared: `COLOR_MAP`, `FONT_MAP`, `hexToNormalizedRgb`, `normalizedRgbMatches`, `driveFileUrl`, `LOGO_CONFIG` |
| `slides-updater.js` | **Modify** | Remove shared items now in `utils.js` (Phase 0, Steps 1–2) |
| `docs-updater.js` | **Create** | All functions from Phase 2 (steps 4–13): `collectDocContent`, `buildDocColorRequests`, `replaceDocColors`, `buildDocFontRequests`, `replaceDocFonts`, `logDocImages`, `buildDocLogoRequests`, `replaceDocLogos`, `updateDocsDocument` |
| `main.js` | **Modify** | Add `updateAllDocsInFolder` (Phase 3, Step 14) |
| `appsscript.json` | **Modify** | Add Advanced Docs API under `dependencies.enabledAdvancedServices` |

---

## Verification

1. After `clasp push`, enable the Advanced Docs API in the Apps Script browser editor (Services panel) and confirm no manifest errors.
2. Confirm `utils.js` globals are accessible: log `COLOR_MAP.length` in `docs-updater.js` — should match the number of color entries.
3. Call `replaceDocColors(testDocId)` on a test document with known old brand colors applied as explicit inline overrides; visually inspect in Google Docs.
4. Confirm that text styled only via Named Styles (no explicit override) is not changed.
5. Call `replaceDocFonts(testDocId)` — confirm Poppins and Figtree text runs are replaced with Lexend.
6. Confirm bold Poppins/Figtree text remains bold after replacement (weight preserved via `weightedFontFamily`).
7. Call `updateDocsDocument(testDocId)` — confirm colors, fonts, and logos all updated in one call.
8. Run `logDocImages(testDocId)` — review Logger output to determine `sourceUri` values and dimensions of existing logos.
9. Set `LOGO_CONFIG.docsLogo.oldSourceUri` based on Step 8 output (if `sourceUri` is populated); otherwise set size bounds.
10. Run `replaceDocLogos(testDocId, true)` (dry run) — check Logger confirms correct logo elements are identified.
11. Adjust `LOGO_CONFIG.docsLogo` if needed, then run `replaceDocLogos(testDocId, false)` — visually confirm logos replaced in body and headers/footers.
12. Call `updateAllDocsInFolder(testFolderId, true)` (dry run) — confirm Logger lists all `.document` MIME type files.
13. Run `updateAllDocsInFolder(testFolderId, false)` on a test folder with 2+ documents — confirm all updated and Logger shows each file processed.
14. Test a document with headers/footers — confirm logos and styled text in those segments are also updated.

---

## Decisions & Scope

- Named Styles are **not** updated (Option A) — Docs REST API has no `updateNamedStyles` request; only explicit inline overrides are changed
- Body, headers/footers, and footnotes are all included (via `collectDocContent`)
- `segmentId` is required on all range requests: `""` for body, header/footer/footnote ID for others
- Logo replacement uses `sourceUri` as primary match (likely populated in Docs); size bounds as fallback
- Logo delete + insert pairs are sorted in **reverse index order** to prevent index shift bugs
- `driveFileUrl` and `LOGO_CONFIG.newLogoFileId` are shared with the Slides updater via `utils.js`
- Scope is batch by Drive folder; MIME type filter is `application/vnd.google-apps.document`
- No undo/rollback mechanism — reversible by running with an inverted color/font map
