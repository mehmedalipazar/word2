# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

SerbestSayfa is a browser-based word processor. Its point is that images are placed freely on pages and never move when the text changes. Word anchors images to paragraphs; this app pins each image to a page instead. Other features:

- Figure captions ("Şekil 1. …") that stay attached to their image.
- `.docx` import and export.
- Print to PDF.

The project has no build step, no dependencies and no `package.json`. The target browsers are Chrome and Edge, the only ones it was tested in. UI strings, toasts and code comments are in Turkish; keep that convention.

## Running and checking

- **Run:** open `index.html` directly over `file://`. All code is classic `<script>` files sharing a global namespace. It must keep working without a server, so don't use ES modules and don't fetch local files.
- **Syntax check:** no linter or test suite exists. Run:
  `Get-ChildItem js\*.js | % { node --check $_.FullName }`
- **Browser automation:** Playwright MCP blocks `file://`. Serve the folder over HTTP, for example `python -m http.server 8765`, then open `http://127.0.0.1:8765/index.html`. In the page, everything is reachable through `SS.app`:
  - `SS.app.state`, `SS.app.layout()`
  - `await SS.importDocx(file)` returns a report; `SS.importSummary(report)` formats it.
  - `await SS.exportDocx(SS.app)` returns a Blob.
- **Print path:** `SS.app.beforePrint()` and `SS.app.afterPrint()` switch the layout into and out of print mode. `page.pdf()` and `chrome --headless=new --print-to-pdf` don't reliably fire `beforeprint`, so call these yourself before and after capturing.
- **Checking `.docx` output:** Word is not installed on this machine; LibreOffice is. Convert with:
  `& "C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to pdf --outdir <dir> <file.docx>`
  LibreOffice breaks lines slightly differently from Chrome and Word. Compare page counts and image positions, not individual lines.
- Keep scratch test output (`.test/`, `.playwright-mcp/`) out of the repo, or delete it afterwards.

## Architecture

Script order in `index.html` is the dependency order:

`util → zip → core → text → objects → docx → docximport → main`

Each file is an IIFE that attaches to `window.SS`. Shared state and cross-module functions live on `SS.app`, which `core.js` creates. Later modules call earlier ones through `app.*`. Core calls back through optional hooks that `main.js` sets:

- `app.onChange` runs after every history commit (autosave, word count, undo buttons).
- `app.onLayout`, `app.onLoad` and `app.onSelectionChange`.

### Layout model (core.js): the central idea

- **Images are not part of the text.** Each image is a plain object:
  `{id, asset, page, x, y, w, h, rot, wrap, locked, z, caption?}`
  `x` and `y` are px (96 DPI) from the top-left of the image's page. The bitmaps live separately in `state.assets[assetId]` as data URLs, which keeps history snapshots small. Units elsewhere: 1 px = 15 twips = 9525 EMU (`SS.units`).
- **All pages share one text flow.** The text is a single `contenteditable` `#editor`. It sits inside `#flow`, which is absolutely positioned at page 0's margin box.
- **No-text zones are float "bands".** Page margins, gaps between pages and text-wrapping images are turned into invisible `.ex` floats:
  - `computeBands()` turns page boundaries and image boxes into horizontal bands. Each band excludes the full width, the left side or the right side.
  - `renderExclusions()` emits each band as `float:left|right; clear:both|right; margin-top:<gap>; shape-outside:border-box`. The bands stack vertically, and their margin area doesn't push text.
- **Flow coordinates:** y=0 is the top of page 0's text area. Page k's text area is `[k*stride, k*stride + ch]`, with `stride = pageHeight + gap`.
- **`app.layout()`** marks empty lines, renders the bands, sizes the page breaks, then recomputes the page count and repeats until the count is stable. The page count comes from where the text ends, which pages have images, and `minPages`.
  - After changing images, call `app.renderImages()` and then `app.layout()`.
  - `app.relayoutAll()` also re-renders the pages.
- **Browser constraints.** Each of these caused a real bug. Don't undo them:
  - Band floats must be **direct children of `#flow`**. Inside a zero-height wrapper, print pagination loses their continuation on later pages.
  - A float must **never be wider than the content width `cw`**. If it is, Chrome also blocks text from its margin-top area.
  - In **print mode** (`app.printMode`, gap = 0), bands are emitted only for real page boundaries and are cut at every physical page boundary. No float may cross a printed page.
  - Lines containing only a `<br>` ignore floats. `markEmptyLines()` adds class `el`, whose CSS `::before{content:"\a0"}` makes those lines skip page gaps too. `serialize()` strips the class.
  - Print CSS sets `orphans`/`widows` to 1, so the browser doesn't re-break lines the bands already placed.
- **Page break:** `<div class="pb" contenteditable="false">`. `fitPageBreaks()` stretches it to the end of the current page's text area on every layout.
- **Captions:** `img.caption = {text, label, pos}`.
  - Rendered as `.cap` next to the image, below or above the rotated image's bounding box. The caption itself is never rotated.
  - Its height is measured from the DOM.
  - `app.objBox(img)` is the image plus its caption. Bands, snapping, alignment and page clamping all use it.
  - Numbers are never stored. `app.captionNumbers()` computes them per label from page, then y, then x order.
- **Layers inside `#doc`:** pages → `#behindLayer` → `#flow` → `#frontLayer` → `#overlay`.
  - Each page has a `.pclip` in both image layers, which clips images to the page.
  - Zoom is a CSS transform on `#doc`. Convert pointer and rect coordinates with `app.toDoc(e)` or `app.clientToFlow()`, which divide by `state.zoom`.

### History and persistence

- **One custom undo stack for text and images.** `app.commit(kind)` snapshots the editor HTML, the images as JSON, and the page settings.
  - Kinds `typing`, `deleting` and `nudge` merge with the previous entry when within 1.5 s.
  - Native undo is intercepted: `beforeinput` `historyUndo`/`historyRedo` and Ctrl+Z.
  - Every user-visible change must end with `app.commit(...)`.
- **`.sayfa` files** are the JSON from `app.serialize()`: `app: 'SerbestSayfa'`, the HTML, the images, the used assets and the page settings.
- **Autosave** goes to the IndexedDB key `autosave`.

### Text editing (text.js)

- Formatting uses `execCommand` with `styleWithCSS`.
- Font size uses the `fontSize 7` marker trick: apply size 7, then rewrite those marked elements to pt.
- Paste goes through `app.sanitizeHTML()`. It reduces Word and web HTML to `p/h1-3/ul/ol/li/b/i/u/s/sub/sup/span/br` and `.pb`, and turns Word's fake list paragraphs into real lists.
- Keep editor HTML within that vocabulary. Export and import only understand it.

### Objects (objects.js)

- Pointer handling runs in the capture phase on `#doc`. The `mousedown` that follows is cancelled through `blockMouse`, so text focus and selection aren't disturbed.
- Images behind the text can be selected only where no glyph is under the pointer (`isOverText`), or with Alt held.
- Captions are edited in place with `contenteditable="plaintext-only"`. Global key and paste handlers skip `.cap-text`.

### DOCX export (docx.js)

- The ZIP is store-only (`zip.js`).
- Images are `wp:anchor` with `relativeFrom="page"`. Each is anchored in the run at the first character of its page.
  - `findPageStarts` finds that character by measuring the live layout, so **export depends on the current DOM layout**.
  - A page with no text gets a holder paragraph with `pageBreakBefore`, and its images are written with `wrapNone`.
- Captions are `wps` text boxes inside `mc:AlternateContent`, with the `Caption` style and a `SEQ` field.
- Page breaks become `pageBreakBefore` on the next paragraph.
- Run and paragraph formatting is read from computed styles.
- Keep OOXML child-element order in `pPr`, `rPr` and `wp:anchor`. Word rejects files with the wrong order.

### DOCX import (docximport.js)

- **Reading:** unzips with `DecompressionStream('deflate-raw')` and resolves styles, theme fonts and numbering. Heading numbers are simulated and written into the text.
- **HTML with placeholders:** drawings become `<span class="ph">` placeholders in the generated HTML.
- **Placement:** after `app.load()`, `placeAll()` goes through the placeholders in document order. For each one it measures the placeholder, creates the image object, and re-runs layout before the next.
  - Inline images become `topbottom` at their line.
  - Anchored images follow Word's positioning rules.
- **Caption attachment** looks for:
  - A paragraph right after (or right before) an image-only paragraph that matches `CAPTION_RE` or uses the `caption` style.
  - A caption-like text box in the same drawing or group.
- **Lossy conversions:** tables become tab-separated paragraphs. EMF/WMF/TIFF images, charts and footnotes are skipped, and `SS.importSummary()` tells the user how many.
