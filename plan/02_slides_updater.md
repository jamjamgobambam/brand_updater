# Plan Part 2: Slides Theme, Color, Font and Logo Updater

## Summary

Update Google Slides presentations by:

1. Replacing the Accent colors in each master slide's `ColorScheme` so all theme-referencing elements update automatically
2. Traversing every page element to replace any hard-coded inline RGB colors that don't use theme references
3. Traversing all text runs to replace Poppins and Figtree fonts with Lexend
4. Replacing logo images on master/layout slides using position heuristics

Uses the **Advanced Slides REST API** (not the basic `SlidesApp` service). A single-file function handles one presentation; a batch wrapper handles multiple from a Drive folder.

---

## Color Mapping

| Accent slot | Old hex | New hex |
|---|---|---|
| Accent 1 | #009eb0 | #003547 |
| Accent 2, Link, and Followed Hyperlink | #9660bf | #005E54 |
| Accent 3 | #ed6060 | #C2BB00 |
| Accent 4 | #3ea33e | #E1523D |
| Accent 5 | #007acc | #ED8B16 |
| Accent 6 | #ead300 | #ead300 (unchanged) |

`DARK1`, `DARK2`, `LIGHT1`, `LIGHT2` — preserved from each master unchanged.

---

## Font Mapping

| Old font | New font |
|---|---|
| Poppins | Lexend |
| Figtree | Lexend |

Text runs with `fontFamily: null` (inheriting from the master placeholder) are left untouched — they pick up the new font automatically once the master's own text runs are updated. Text runs with any other explicit font are also left untouched.

---

## Logo Config

Logos are identified by **position heuristics** — where the center of an image element falls as a percentage of the slide dimensions. Two logo types are matched:

| Logo type | Position rule |
|---|---|
| Corner logo (recurring, bottom-right) | `centerX > 75%` of slide width AND `centerY > 75%` of slide height |
| Title logo (title slide, upper-center) | `centerX` between 25–75% AND `centerY < 35%` |

Because logos live on **master/layout slides**, each logo is replaced once there and all individual slides inherit the change automatically.

The new logo must be a Google Drive file shared as "Anyone with the link can view". Its Drive file ID is stored in `LOGO_CONFIG` and converted to a direct URL at runtime.

---

## Phase 1 — Manifest Prerequisites

### Steps

1. Add to `appsscript.json` under `dependencies.enabledAdvancedServices`:

   ```json
   { "userSymbol": "Slides", "serviceId": "slides", "version": "v1" }
   ```

   Also enable it in the cloud project via the Apps Script editor → Services panel.

The basic `SlidesApp` service (Apps Script's built-in service) gives you a high-level object model but **cannot read or write the `ColorScheme`** on a master slide. That requires the full REST API, exposed in Apps Script as an "Advanced Service." It runs server-side within Apps Script without needing manual OAuth tokens.

Apps Script's manifest declares which advanced services your script uses. If it's not listed there, the `Slides` global won't exist at runtime. Google enforces a two-step requirement (manifest + Services panel) to prevent accidental API usage.

---

## Phase 2 — File Structure

### Steps

1. Create `utils.js` — shared constants and helpers used by both Slides and Docs updaters
2. Create `slides-updater.js` — all Slides feature logic (references globals from `utils.js`)
3. Create `main.js` — batch entry point + public trigger functions

Apps Script treats all `.js` files in the project as a single global scope (they're concatenated at runtime). Splitting by feature keeps things navigable — `slides-updater.js` is self-contained and later you can add `docs-updater.js` without touching the Slides logic. `main.js` is the "entry point" layer — the functions you'd actually click "Run" on in the editor, like `updateAllSlidesInFolder`.

**`utils.js` contains:** `COLOR_MAP`, `FONT_MAP`, `hexToNormalizedRgb`, `normalizedRgbMatches`, `driveFileUrl`, `LOGO_CONFIG`, and `getPresentation`. These are defined once here and available to both `slides-updater.js` and `docs-updater.js` via the shared global scope — no import statements needed. Do **not** redefine these in `slides-updater.js` or `docs-updater.js`; duplicate `const` declarations across files cause a runtime error.

---

## Phase 3 — Implement `slides-updater.js`

### Step 4 — `COLOR_MAP` constant (in `utils.js`)

An array of `{ oldHex, newHex }` pairs (6 entries from the color table above). Also defines the target hex for `HYPERLINK` and `FOLLOWED_HYPERLINK` (both `#005E54`) used in Step 7.

This serves as a single source of truth for the mapping. Every function that needs to know "what replaces what" reads from here — if you later add a color or change a target, you update one place.

---

### Step 5 — `hexToNormalizedRgb(hex)`

Converts `#RRGGBB` to `{ red, green, blue }` in the 0.0–1.0 range.

The Slides API doesn't use 0–255 integers or hex strings for colors — it uses floats between 0.0 and 1.0. So `#009eb0` becomes `{ red: 0/255, green: 158/255, blue: 176/255 }` = `{ red: 0.0, green: 0.6196..., blue: 0.6902... }`. Every place you build a request or compare a color needs this conversion.

---

### Step 6 — `normalizedRgbMatches(apiRgb, targetHex, tolerance)`

Compares an API `rgbColor` object against a target hex with float tolerance (~0.004 = 1/255). Returns boolean.

When the API returns a stored color like `green: 0.6196078431372549`, and you compute `158/255 = 0.6196078431372549`, they *might* match exactly — but floating-point storage and rounding in Google's backend can introduce tiny errors. A tolerance of `1/255 ≈ 0.00392` is the smallest meaningful color unit, so anything within that is the same color. Without this, you'd miss valid matches.

---

### Step 7 — `updateMasterThemeColors(presentationId, masters)`

*Depends on steps 4, 5.*

For each master:

- Read its existing `pageProperties.colorScheme.colors` array
- Deep-copy it and replace only the `ACCENT1`–`ACCENT6`, `HYPERLINK`, and `FOLLOWED_HYPERLINK` type entries using `COLOR_MAP`
- Build one `updatePageProperties` request per master with `fields: "colorScheme"`
- Submit via `Slides.Presentations.batchUpdate(presentationId, { requests })`

Master slides contain the `ColorScheme` — an ordered array of named color entries like `{ type: "ACCENT1", color: { rgbColor: {...} } }`. When you change these, every element on every layout and slide that *references* a theme color (e.g., "use Accent 1 for this shape fill") automatically reflects the new color — Google Slides resolves these references at render time.

**You must read the existing scheme first and patch only the entries you're changing**, then write the whole thing back using `fields: "colorScheme"`. If you omit `DARK1` or `LIGHT2` from your write, the API treats them as "unset" and may blank them. Reading first → patching → writing back is the safe pattern.

---

### Step 8 — `buildInlineColorRequests(pages, colorMap)`

*Depends on steps 4, 5, 6.*

Core traversal function. Iterates all pages — checking both page-level properties and the `pageElements` of each page — and builds the corresponding request type for each color location found:

| Location | API path | Request type |
|---|---|---|
| Page background | `pageProperties.pageBackgroundFill.solidFill.color.rgbColor` | `updatePageProperties` |
| Shape fill | `shapeProperties.shapeBackgroundFill.solidFill.color.rgbColor` | `updateShapeProperties` |
| Shape outline | `shapeProperties.outline.outlineFill.solidFill.color.rgbColor` | `updateShapeProperties` |
| Text runs | `shape.text.textElements[].textRun.style.foregroundColor.rgbColor` | `updateTextStyle` (with `startIndex`, `endIndex`, `objectId`) |
| Table cells | `table.tableRows[].tableCells[].tableCellProperties.tableCellBackgroundFill.solidFill.color.rgbColor` | `updateTableCellProperties` |
| Line fill | `line.lineProperties.lineFill.solidFill.color.rgbColor` | `updateLineProperties` |

Returns array of batchUpdate request objects.

Not every color in a presentation uses the theme. Sometimes we pick a color directly from the color picker — this stores an explicit `rgbColor` on the element (an "inline" or "direct" color). These are **not** affected by changing the master `ColorScheme`, so they need to be found and replaced individually.

Page background is a property of the page itself (not any element within it), so it requires its own `updatePageProperties` request targeting the page's `objectId` — it is not reachable via the per-element traversal used for shape fills and text runs.

Each location maps to a different `batchUpdate` request type because the API distinguishes between updating a shape's properties vs. a text range vs. a table cell.

For text runs specifically, you need `startIndex`/`endIndex` because a single text box might have mixed styles — you can't replace the whole text box's color, only the specific run that has the matching color.

---

### Step 9 — `replaceInlineColors(presentationId)`

*Depends on step 8.*

- Calls `Slides.Presentations.get(presentationId)` to get full presentation JSON
- Collects all page categories: `presentation.masters`, `presentation.layouts`, `presentation.slides`
- Calls `buildInlineColorRequests(allPages, COLOR_MAP)`
- If requests are non-empty, submits via `Slides.Presentations.batchUpdate()`

This gets the full presentation JSON in **one API call** — important because you want one read, not one per slide. Masters and layouts can also have inline-colored elements (e.g., a logo or decorative element on the layout template), so all three page categories must be included.

**batchUpdate request limit:** The Slides API allows a maximum of ~500 requests per `batchUpdate` call. A presentation with many inline-colored text runs could approach this limit. The implementation always proactively splits requests into chunks of 400 and submits multiple sequential `batchUpdate` calls — this is the default behavior, not a fallback triggered by errors.

---

### Step 10 — `updateSlidesPresentation(presentationId, dryRun)`

Public orchestrator. Sequences:

1. `Slides.Presentations.get()` to read masters
2. `updateMasterThemeColors()`
3. `replaceInlineColors()`
4. `replaceFonts()`
5. `replaceLogos(presentationId, dryRun)`

Kept separate from steps 7, 9, 13, and 19 for testability and reuse — each function can be called and tested independently. Colors run first (theme palette correct before inline cleanup), then fonts, then logos. The `dryRun` flag is passed through to `replaceLogos` so you can safely audit logo matches before committing.

---

### Step 11 — `FONT_MAP` constant (in `utils.js`)

An array of `{ oldFont, newFont }` pairs:

- `{ oldFont: "Poppins", newFont: "Lexend" }`
- `{ oldFont: "Figtree", newFont: "Lexend" }`

Same single-source-of-truth pattern as `COLOR_MAP`. Add or change a font mapping in one place.

---

### Step 12 — `buildFontRequests(pages, fontMap)`

Traversal identical in structure to `buildInlineColorRequests` — same page collection (masters + layouts + slides), same element iteration, same text run path. For each text run:

- Read `textRun.style.weightedFontFamily.fontFamily` (preferred — includes weight)
- Fall back to `textRun.style.fontFamily` if `weightedFontFamily` is absent
- If either matches an old font name, build an `updateTextStyle` request with `weightedFontFamily: { fontFamily: newFont, weight: <existing weight> }` and `fields: "weightedFontFamily"`

Returns array of `updateTextStyle` request objects.

There is no central `FontScheme` on the master equivalent to `ColorScheme` — fonts are stored per text run, so every run with an explicit matching font must be updated individually.

The `weightedFontFamily` detail is critical: the API stores font and weight (e.g. 400, 700) together. Writing only `fontFamily` silently resets the weight to normal — bold text becomes regular. Preserving the existing weight in the request prevents this.

Text runs where `fontFamily` is `null` are never matched, which is intentional — they inherit the new font from the updated master automatically.

---

### Step 13 — `replaceFonts(presentationId)`

*Depends on step 12.*

- Calls `getPresentation(presentationId)` to get full presentation JSON (with retry — see `utils.js`)
- Collects all page categories: `presentation.masters`, `presentation.layouts`, `presentation.slides`
- Also collects **speaker notes pages** by reading `slide.slideProperties.notesPage` for each slide and appending it to the pages array
- Calls `buildFontRequests(allPages, FONT_MAP)`
- If requests are non-empty, submits via `Slides.Presentations.batchUpdate()`

Same single-read pattern as `replaceInlineColors` — one API call for the full JSON, then process all page categories. Masters and layouts must be included because placeholder text on those pages also has explicit font settings that slide-level changes won't override. Speaker notes pages are included so that font replacements apply to the notes text as well as the slide content — the `notesPage` object has the same `pageElements` structure, so `buildFontRequests` handles it without any changes.

---

### Step 14 — `logAllImages(presentationId)`

Diagnostic/discovery utility. Run this **once on a representative presentation before tuning `LOGO_CONFIG`**.

- Calls `Slides.Presentations.get(presentationId)` for `pageSize`, `masters`, and `layouts`
- Iterates every image element on those pages
- For each image, logs:
  - `objectId` and page name
  - `centerX` and `centerY` as computed percentages (0.0–1.0)
  - Width and height as percentages of slide dimensions
  - `sourceUrl` if present (non-empty means URL matching is possible as a future improvement)
- Does not make any changes

**Why:** With position heuristics there's no way to know if the default thresholds in `LOGO_CONFIG` will correctly identify logos without first seeing where actual logo images land in the coordinate space. Running this on a real presentation tells you exactly what values to use for `xThreshold`, `yThreshold`, `xMin`, `xMax`, and `yMax` before touching any files. It also reveals whether `sourceUrl` is populated, which would allow switching to more reliable URL matching later.

---

### Step 15 — `LOGO_CONFIG` constant (in `utils.js`)

```js
const LOGO_CONFIG = {
  newLogoFileId: "YOUR_DRIVE_FILE_ID",
  cornerLogo: { xThreshold: 0.75, yThreshold: 0.75 },
  titleLogo:  { xMin: 0.25, xMax: 0.75, yMax: 0.35 }
};
```

Thresholds are percentages of the slide dimensions. Keeping them in a constant means you can tune the detection zone without touching any logic code.

---

### Step 16 — `driveFileUrl(fileId)` (in `utils.js`)

Builds `https://drive.google.com/uc?export=download&id=${fileId}`.

This helper is defined in `utils.js` as a shared utility. It is **not used by the Slides logo replacement** (see Step 19 — the implementation switched to a Drive blob approach). It is retained for potential future use and is used by the Docs updater as a fallback URL format for logo insertion.

---

### `getPresentation(presentationId, maxAttempts)` (in `utils.js`)

Wraps `Slides.Presentations.get()` with retry logic for transient API failures (e.g. `"Empty response"` errors that occasionally occur on valid presentation IDs).

- Attempts the call up to `maxAttempts` times (default: 3)
- On failure, waits `2^attempt` seconds before retrying (1 s, then 2 s)
- Re-throws the error if all attempts fail

Used in place of bare `Slides.Presentations.get()` calls in `replaceFonts`. The Slides API occasionally returns an empty response for a valid, accessible presentation — a single retry is almost always sufficient to recover. Using exponential backoff avoids hammering the API immediately after a transient failure.

---

### Step 17 — `isLogoElement(element, pageWidth, pageHeight, type)`

- Returns `false` immediately if `element.image` is absent (non-image elements are skipped)
- Returns `false` immediately if `element.transform` is absent — elements positioned at their default location have no explicit `transform` object in the API response; accessing `element.transform.translateX` without this guard would throw a TypeError at runtime
- Reads `element.transform` (translateX, translateY) and `element.size` (width, height) from the element's `AffineTransform` — all values in EMUs
- Computes `centerX = (translateX + width / 2) / pageWidth` and `centerY = (translateY + height / 2) / pageHeight`
- Checks computed percentages against `LOGO_CONFIG.cornerLogo` or `LOGO_CONFIG.titleLogo` thresholds depending on `type`
- Returns boolean

EMUs (English Metric Units) are the raw unit the Slides API uses for all positions and sizes. There are 914,400 EMUs per inch and 12,700 EMUs per point. The slide's `pageSize` is also in EMUs, so dividing by it gives a clean 0.0–1.0 percentage.

---

### Step 18 — `buildLogoReplaceRequests(pages, pageWidth, pageHeight, newLogoUrl, dryRun)`

- Iterates masters and layouts only (individual slides inherit — no need to traverse them)
- For each element, calls `isLogoElement()` for both `"corner"` and `"title"` types
- **Known limitation:** grouped elements (`element.elementGroup`) are not recursed into — if a logo is nested inside a group, it will be skipped. Use `logAllImages` output to verify; if no image appears at the expected position, a group wrapper may be the cause
- On a match:
  - If `dryRun` is `true`: logs the element's objectId, page name, and detected type — no request built
  - If `dryRun` is `false`: builds a `replaceImage` request with `imageReplaceMethod: "CENTER_INSIDE"` and the new logo URL
- Returns array of `replaceImage` request objects (empty if `dryRun`)

The dry-run mode is critical for position-heuristic matching because there's no perfectly deterministic way to identify logos without a content URL. Running with `dryRun: true` first lets you review what the script would replace — and adjust the thresholds in `LOGO_CONFIG` if needed — before making any changes.

`CENTER_INSIDE` preserves the original bounding box dimensions and centers the new image within it, which is the correct behavior when swapping logos of similar but not identical aspect ratios.

---

### Step 19 — `replaceLogos(presentationId, dryRun)`

*Depends on steps 15, 17, 18.*

- Calls `Slides.Presentations.get(presentationId)` to get presentation JSON including `pageSize`, `masters`, and `layouts`
- Extracts `pageWidth` and `pageHeight` from `presentation.pageSize`
- **Dry-run path:** calls `buildLogoReplaceRequests(..., null, true)` for logging only, then returns
- **Live path:**
  1. Calls `buildLogoReplaceRequests(..., "_placeholder_", false)` — the placeholder URL is never used; this call is only to collect the `objectId`s of matched logo elements
  2. Builds a `Set` of logo `objectId`s from the returned request objects
  3. Fetches the new logo as a blob: `DriveApp.getFileById(LOGO_CONFIG.newLogoFileId).getBlob()`
  4. Opens the presentation via `SlidesApp.openById(presentationId)`
  5. Iterates `deck.getMasters()` and `deck.getLayouts()`; for each image whose `objectId` is in the set, calls `img.replace(logoBlob, false)`

**Why blob instead of REST `replaceImage` + URL:** The initial plan used `Slides.Presentations.batchUpdate` with a `replaceImage` REST request and a public Drive URL. In practice this approach is unreliable — the Slides REST `replaceImage` API requires the URL to be directly accessible without redirects, and Google Drive URLs (even with `export=download`) can trigger anti-abuse redirects that cause silent failures. Using `SlidesApp.Image.replace(blobSource, crop=false)` fetches the image bytes via Drive's internal service, requires no public sharing on the logo file, and produces the same scale-to-fit (CENTER_INSIDE equivalent) behavior.

**No public sharing required:** The new logo Drive file does **not** need to be publicly shared. `DriveApp.getFileById()` reads the file with the script owner's credentials.

Only masters and layouts are iterated — individual slides inherit the change automatically.

---

## Phase 4 — Batch Wrapper in `main.js`

### Step 20 — `updateAllSlidesInFolder(folderId, dryRun)`

Entry point for batch runs:

- Uses `DriveApp.getFolderById(folderId).getFiles()` to iterate files
- Filters by MIME type `application/vnd.google-apps.presentation`
- Wraps each `updateSlidesPresentation(file.getId(), dryRun)` call in a `try/catch` — logs the file name and error message on failure and continues to the next file rather than aborting the entire batch
- Logs progress (file name, success or failure) with `Logger.log()`

The Drive API lets you list files in a folder and filter by MIME type. `application/vnd.google-apps.presentation` is Google's MIME type for Slides files. Logging with `Logger.log()` is essential here because Apps Script has no console — the Execution Log in the browser editor is how you see what ran and whether anything failed. For a batch job touching many files, you want a trace of each file processed.

**Apps Script execution timeout:** Apps Script hard-kills execution after 6 minutes (30 minutes for Google Workspace accounts). A large folder could exceed this limit. If that happens, run the updater on smaller subfolders rather than the root folder. A more robust future improvement would be to store a file-index cursor in `PropertiesService` so a subsequent run can resume from where the previous one stopped.

---

## Files to Create / Modify

| File | Action | Purpose |
|---|---|---|
| `appsscript.json` | Modify | Add Advanced Slides API under `dependencies.enabledAdvancedServices` |
| `utils.js` | **Create first** | Shared: `COLOR_MAP`, `FONT_MAP`, `hexToNormalizedRgb`, `normalizedRgbMatches`, `driveFileUrl`, `LOGO_CONFIG`, `getPresentation` |
| `slides-updater.js` | Create | Logic functions from Phase 3 (colors: steps 5–9, fonts: steps 12–13, logos: steps 14, 17–19) — references globals from `utils.js` |
| `main.js` | Create | Batch wrapper (Phase 4) + trigger entry points |

---

## Verification

1. After `clasp push`, enable the Advanced Slides API in the Apps Script browser editor (Services panel) and confirm no manifest errors.
2. Log-test `hexToNormalizedRgb("#009eb0")` — expect `{ red: 0, green: ~0.6196, blue: ~0.6902 }`.
3. Call `updateSlidesPresentation(testId)` on a test presentation with known old accent colors; visually inspect in Google Slides.
4. Verify theme-referenced elements updated (color swatch in the Slides theme editor shows new palette).
5. Verify inline-colored elements (direct RGB, not theme-referenced) also updated.
6. Call `updateAllSlidesInFolder(testFolderId)` with a folder containing 2+ test presentations; confirm all updated and Logger output shows each file processed.
7. Confirm Poppins and Figtree text has been replaced with Lexend across masters, layouts, and slides.
8. Confirm bold Poppins/Figtree text is still bold after replacement (weight preserved via `weightedFontFamily`).
9. Confirm text with a different explicit font (e.g. "Roboto") is unchanged.
10. Confirm Poppins and Figtree text in speaker notes has also been replaced with Lexend.
10. Confirm runs with no explicit font set (inheriting) still render correctly after the master update.
11. Run `logAllImages(testId)` on a representative presentation — review Logger output to determine the correct position percentage thresholds for logos and whether `sourceUrl` is populated for any images.
12. Set `LOGO_CONFIG` thresholds based on the output of step 11.
13. Run `replaceLogos(testId, true)` (dry run) — check Logger output confirms the right image elements are identified as corner/title logos and nothing unexpected is matched.
14. Adjust `LOGO_CONFIG` thresholds if needed, then run `replaceLogos(testId, false)` — visually confirm the corner logo on layout slides and the title logo on the title layout are replaced.
15. Confirm position and size are preserved (new logo in the same bounding box).
16. Confirm the Drive file is set to "Anyone with the link can view" — if not, `replaceImage` will throw a URL access error.

---

## Decisions & Scope

- `FOLLOWED_HYPERLINK` → `#005E54` (same as Link/`HYPERLINK`)
- `DARK1`, `DARK2`, `LIGHT1`, `LIGHT2` are **not** changed (preserved from current master)
- Accent 6 → `#ead300` maps to itself; still explicitly set for completeness
- Scope is Google Slides only; Google Docs is a separate plan
- No undo/rollback mechanism — reversible by running with an inverted color/font map
- Null/inheriting `fontFamily` runs are intentionally not matched — they pick up the new font from the master automatically
- Font weight is preserved via `weightedFontFamily` to prevent bold text from losing its weight on replacement
- Logo replacement uses position heuristics — always run with `dryRun: true` first to audit matches before committing
- Logo replacement uses `SlidesApp.Image.replace(blob, false)` which scales to fit the existing bounding box (equivalent to CENTER_INSIDE) — the Drive file does **not** need to be publicly shared
- Only masters and layouts are traversed for logos — individual slides inherit the change automatically. Any logos placed directly on individual slides (not via master/layout inheritance) will not be replaced and require manual cleanup.
