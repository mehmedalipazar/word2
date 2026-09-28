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

`util → zip → core → text → find → ruler → objects → docx → docximport → main`

Each file is an IIFE that attaches to `window.SS`. Shared state and cross-module functions live on `SS.app`, which `core.js` creates. Later modules call earlier ones through `app.*`. Core calls back through optional hooks that `main.js` sets:

- `app.onChange` runs after every history commit (autosave, word count, undo buttons).
- `app.onLayout`, `app.onLoad` and `app.onSelectionChange`.
- `app.onStyles` (set by `text.js`) runs after `applyStyles()` (load, undo, style changes) and syncs the heading numbering box.

### Layout model (core.js): the central idea

- **Images are not part of the text.** Each image is a plain object:
  `{id, asset, page, x, y, w, h, rot, wrap, locked, z, caption?, crop?}`
  `x` and `y` are px (96 DPI) from the top-left of the image's page. The bitmaps live separately in `state.assets[assetId]` as data URLs, which keeps history snapshots small. Units elsewhere: 1 px = 15 twips = 9525 EMU (`SS.units`).
  - `crop = [left, top, right, bottom]` are the fractions (0–1) cut from the original bitmap; `x/y/w/h` is the visible frame. Cropping never changes the asset: `.img-obj.cropped` clips and the inner `<img>` is offset with `app.cropStyle(crop)`. Export writes `a:srcRect`; import keeps Word's `srcRect` as `crop` (no longer bakes it into the bitmap). Always assign a new array (clones share it).
- **All pages share one text flow.** The text is a single `contenteditable` `#editor`. It sits inside `#flow`, which is absolutely positioned at page 0's margin box.
- **No-text zones are float "bands".** Page margins, gaps between pages and text-wrapping images are turned into invisible `.ex` floats:
  - `computeBands()` turns page boundaries and image boxes into horizontal bands. Each band excludes the full width, the left side or the right side.
  - `renderExclusions()` emits each band as `float:left|right; clear:both|right; margin-top:<gap>; shape-outside:border-box`. The bands stack vertically, and their margin area doesn't push text.
- **Flow coordinates:** y=0 is the top of page 0's text area. Page k's text area is `[k*stride, k*stride + ch]`, with `stride = pageHeight + gap`.
- **`app.layout()`** marks empty lines, renders the bands, sizes the page breaks, then recomputes the page count and repeats until the count is stable. The page count comes from where the text ends, which pages have images, and `minPages`.
  - After changing images, call `app.renderImages()` and then `app.layout()`.
  - `app.relayoutAll()` also re-renders the pages.
- **Pagination rules (Word's widow/orphan control, keep-with-next and keep-lines for headings).** CSS `orphans`/`widows` only work in real fragmentation, so they are emulated with bands:
  - `keep[k]` shortens page k's text area from the bottom: `computeBands()` adds a full band `[k*stride + ch - keep[k], k*stride + ch]`, which pushes the offending lines to the next page. Print mode uses the same bands.
  - `paginateBoundary(k)` decides one page boundary from the live layout (`boundaryCut`: measure line rects of the paragraph units around the boundary). Boundaries are processed in order; every band change is a full relayout (~28 ms for 47 pages), so the work is incremental:
    - `app.markDirty(from, to)` records changed pages. Text edits mark the selection's pages (`markDirtyAtSelection`, measured inside the next `layout()`), `renderImages()` diffs image signatures and marks their old/new pages, load/undo/page setup mark everything.
    - The pass stops early once a page past the dirty range starts at the same line as before (`pageKey`).
    - While typing it runs in the background (300 ms after the last change, ≤24 ms chunks). `app.paginateNow()` finishes it synchronously: export, print (`beforePrint`) and docx import (before measuring each placeholder, up to its page) call it.
    - History snapshots store `keep`, so undo/redo lay out correctly at once.
  - docx export writes the same rules (`widowControl` in `pPrDefault`, `keepNext`/`keepLines` in heading styles) so page breaks match Word.
- **Editing commands and bands.** Chrome's editing commands relayout at every step; with dozens of band floats this grows quadratically (bold on 47 pages took 14 s). `app.withoutBands(fn)` runs a command with the bands hidden (`.flow.nobands`), and `beforeinput` does the same for native edits of multi-paragraph selections. Only 3 bands are pre-built past the last page (`EXTRA_BANDS`).
- **Browser constraints.** Each of these caused a real bug. Don't undo them:
  - Band floats must be **direct children of `#flow`**. Inside a zero-height wrapper, print pagination loses their continuation on later pages.
  - A float must **never be wider than the content width `cw`**. If it is, Chrome also blocks text from its margin-top area.
  - In **print mode** (`app.printMode`), bands are emitted only for real page boundaries and are cut at every physical page boundary. No float may cross a printed page.
  - **Print pitch.** Chrome breaks printed pages every `ceil(page height snapped to 1/64 px)` CSS px, not every page height (A4 29.7 cm = 1122.52 px → 1123; Letter 1056 → 1056; `printPitch` in main.js). `beforePrint` therefore sets the gap to pitch − page height, so pages, bands and cuts sit on the real breaks; the extra fraction lies below each page and isn't printed. With gap 0 every page drifted ~0.5 px: page starts crept into the top margin, pushes across a page boundary landed before the real break, and the end of long documents was clipped off the last page.
  - **Band edges are snapped to 1/64 px** (Chrome's layout unit) in `drawBands` before the `margin-top` chain is built. Otherwise each float's rounding adds up along the chain (up to 1/32 px per page: text jumped 1 px halfway through long documents, on screen and in print).
  - Print CSS doesn't clip `#doc`/`#flow` vertically: if printed pagination ever diverges from the layout, text that runs past the last page prints on an extra page instead of vanishing.
  - Lines containing only a `<br>` ignore floats. `markEmptyLines()` adds class `el`, whose CSS `::before{content:"\a0"}` makes those lines skip page gaps too. `serialize()` strips the class.
  - Print CSS sets `orphans`/`widows` to 1, so the browser doesn't re-break lines the bands already placed.
  - List markers must stay `list-style-position: inside`. With band floats present, Chrome loses `outside` markers (the floats' margin-top area breaks their placement). The hanging indent comes from `li { padding-left: var(--li-pad); text-indent: var(--li-ti) }` plus a tab at the end of `::marker` content that aligns the text to the hanging indent (`tab-size: max(1.27cm, hanging)`; see List indent).
  - **A first line with negative `text-indent` next to a float** (list item, numbered heading, hanging paragraph) starts |text-indent| back from the float's edge. Beside a full-width band (page gap, pagination keep band, top-bottom image) that leaves a sliver at the right edge, where the number, marker or first word lands on the previous page. `pushHanging` moves such blocks past the band with `margin-top` rules in the layout's own stylesheet `<style id="hangFix">` (selected by `nth-child` path, rebuilt on every band render, never in the document HTML). `pushState` keeps each pushed block's natural margin, because its unpushed position can't be measured any more. Bands are decided in document order, and a band's rules are written before the next band is measured: a push shifts the following pages, and rules computed from one measurement never settled in long documents (stale pushes left blank gaps mid-page). A pushed block already sitting just after its band keeps its rule, otherwise sub-pixel noise rewrites the sheet (and forces a relayout) on every layout. Rules are re-targeted first, because inserted or removed blocks shift the `nth-child` paths. One line added near the top can change the push at every page gap, so `layout()` settles at most `HANG_BUDGET` bands and pagination does the rest: `paginateBoundary(k)` settles all bands up to boundary k before measuring it, the pass doesn't stop early while pushes are pending, `paginationPending()` includes them, and `paginateNow()` and print mode settle all. Don't replace this with padding. In real printing (fragmented layout), Chrome won't split block-start padding at a physical page boundary: it moves the box with its padding to the next page, or doesn't push it at all. In print mode a push across a page boundary therefore puts the box on the boundary with `margin-top` (margins at unforced breaks are dropped) and pads only the top margin. Verify such changes in a real PDF (`page.pdf`), not only with emulated print media.
  - Beside a left image band, Word/LibreOffice measure a block's indents from the image edge. `indentRows` widens the band along each block by the block's left indent (`computeBands(extra)`), so markers and hanging first lines don't go under the image. `renderExclusions` settles pushes and widened rows within a few rounds.
- **Page break:** `<div class="pb" contenteditable="false">`, always a direct child of `#editor`. `fitPageBreaks()` stretches it to the end of the current page's text area on every layout. `normalizeBlocks` hoists any `.pb` found inside another block (Ctrl+Enter in a list item, old files) with `app.splitTop()`, which splits the top-level block at a position (lists keep numbering via `ol[start]`). A `.pb` inside a list item made background pagination add pages without end and froze the tab.
- **Styles** (`state.styles = {p, h1, h2, h3, num}`: Word's Normal and Heading 1–3; font, size pt, bold, italic, color, align, before/after pt, line = Word multiple; `num` = heading numbering, below). `app.applyStyles()` writes the editor's CSS for them into `<style id="docStyles">` (after `css/app.css`, whose defaults equal the built-in styles; Word's "1.08" is exactly 259/240, i.e. `line-height: 1.3177` in Calibri). Headings get their own font's line factor. Header/footer and captions use Normal's font. `app.setStyles()` = apply + relayout + one history step. Styles are in `.sayfa`, history snapshots and autosave; files without styles load with the built-in ones (not the user's defaults). `app.defaultStyles()` = "Varsayılan olarak ayarla" (localStorage) or built-in, used for new documents and the first blank one.
- **Heading numbering** (Word's multilevel list linked to Heading 1–3): `styles.num` is `null` or `[L1, L2, L3]`, each `null` (level not numbered) or `{text, fmt, start, ind, hang, suff, lgl}`. `text` is Word's lvlText (`%1`–`%3` = the Heading 1–3 counters, only own and higher levels, checked by `cleanStyles`), `ind`/`hang` in pt.
  - `headingNumCSS` draws the numbers with CSS counters `hn1`–`hn3` reset on `.editor`. Heading k increments its counter and resets the lower ones with **`counter-set`**: a `counter-reset` on sibling headings does not reset the next section ("2.3.").
  - The number is `::before` content, so it is not in the text, the word count, Find or the clipboard. With a tab suffix it is an inline-block with `min-width: hang`, so the text starts at the hanging indent like Word.
  - `h1–h3[data-num="0"]` is an unnumbered heading (Word `numId 0`). `app.isNumbered(block)` tells whether a heading shows a number.
  - `app.setIndent(block, prop, px)` writes an indent relative to the stylesheet value (the numbering indent): `''` means "back to the numbering", and 0 can still be set. The ruler, the Paragraph dialog and Ctrl+M use it.
  - The toolbar box `#headingNum` applies the `app.HEADING_NUMS` presets and toggles `data-num`.
- **Tab stops.** A block's custom stops are `data-tabs`: `l|c|r|d` + position in px from the text area's left edge + optional leader `.`/`-`/`_`, e.g. `l75.6 r604.73.` (`app.parseTabs`/`app.formatTabs`).
  - In such blocks every tab character is its own `span.tab` with `tab-size: 0` (`normalizeBlocks` → `wrapTabs`, which also unwraps spans that got typed text and keeps the caret).
  - `fitTabs()` runs in every `layout()`, cached per block signature. The signature includes the image side bands beside the block (`sideBands` from `renderExclusions`), because a wrapping image shifts where lines start. It resets all of a block's tabs to 0, then gives each tab `padding-left` up to its stop, left to right: left, right/center (the following segment up to the next tab, minus 0.2 px so rounding can't break the line), decimal (up to the first `,`, Turkish decimal separator; without a comma the number ends at the stop, `.` being a thousands separator; without a number like a right tab). Past the last stop the default 1.27 cm stops apply, and a hanging indent acts as a stop. No default stop beyond the right margin: the tab gets 0 and the text after it wraps.
  - Leaders are CSS backgrounds (`.lead-*`). `serialize()` strips the paddings. Blocks without `data-tabs` keep CSS `tab-size: 1.27cm`.
  - Ruler: the type box `#rulerTabType` picks the stop type; a click adds a stop, dragging moves it, dragging it off the ruler removes it, and a double-click opens `app.tabsDialog()`, which is also "Sekmeler…" in the Paragraph dialog.
- **List formats** (Word's numbering/bullet library and multilevel lists): a list's `data-lf` is `level1;level2;…`, each level `fmt|text`.
  - `fmt` is Word's `numFmt` (the `app.NUM_FMTS` keys) or `bullet`. `text` is Word's `lvlText`, with `%1` = the list's own items and `%2` = one list deeper; for a bullet, the character itself.
  - Level 1 applies to the list's own items; level k applies to lists k−1 deep that have no `data-lf` of their own, when the type fits (numbered → `ol`, bullet → `ul`). Missing levels use the defaults in `css/app.css` (1. a. i. / • o ▪).
  - `listFmtCSS` generates the `::marker` rules into `<style id="listStyles">`: `counter(list-item, …)`, or `counters(list-item, ".")` for `%1.%2.`. `normalizeBlocks` → `applyListStyles` validates the attribute (`app.parseLF`/`formatLF`) and adds rules for new values.
  - Adjacent lists are separate numbering units. Import and the app's own splits put `margin-bottom: 0` / `margin-top: 0` inline between lists that were one block in Word (same style with `contextualSpacing`) or in the editor.
  - UI: the `#listFmt` box (`app.setListFormat`: writes the level on the owner list = nearest ancestor with `data-lf`, else the top list; retags `ol`↔`ul`; a list Chrome merged into an adjacent list of another format is split off). "Numaralandırma değerini ayarla" (`#numValueDialog`) sets `ol[start]`, splitting the list at a middle item.
  - Chrome doesn't renumber the following items after `indent`/`outdent` or list commands in the `ol > ol` form ("1. 2. a. 4."); `refreshLists()` toggles `display` on the affected top-level lists.
- **Captions:** `img.caption = {text, label, pos, chapter?}`.
  - Rendered as `.cap` next to the image, below or above the rotated image's bounding box. The caption itself is never rotated.
  - Its height is measured from the DOM.
  - `app.objBox(img)` is the image plus its caption. Bands, snapping, alignment and page clamping all use it.
  - Numbers are never stored. `app.captionNumbers()` computes them per label from page, then y, then x order, and returns `id → {n, ch, sep, text}`.
  - `caption.chapter` (a separator `.` `-` `–` `—` `:`) is Word's "include chapter number": the number is the last Heading 1 above the image (`chapterMarks`: its page and y compared with the image's; the heading numbering counter and format, or the heading order when headings aren't numbered), the separator, and the sequence, which restarts at each Heading 1. `refreshChapterCaptions()` renumbers after every `layout()`. The context bar's `#capNum` sets it for all captions of the label.
- **Layers inside `#doc`:** pages → `#behindLayer` → `#flow` → `#frontLayer` → `#overlay`.
  - Each page has a `.pclip` in both image layers, which clips images to the page. `renderPages()` keeps the existing clips (it only resizes, adds or removes them), and `renderImages()` moves an image only when it is in the wrong clip or out of z-order. Detaching or moving a caption being edited blurs it and ends the edit: letters typed after a page-count change were lost.
  - Zoom is a CSS transform on `#doc`. Convert pointer and rect coordinates with `app.toDoc(e)` or `app.clientToFlow()`, which divide by `state.zoom`.

### History and persistence

- **One custom undo stack for text and images.** `app.commit(kind)` snapshots the editor HTML, the images as JSON, the page settings (incl. header/footer and styles) and the pagination bands (`keep`).
  - Kinds `typing`, `deleting` and `nudge` merge with the previous entry when within 1.5 s and the caret hasn't moved since (typing somewhere else starts a new step).
  - Each entry keeps two selections: `sel` (after the change) and `before` (before it). Undo restores the previous state with the undone entry's `before` (Word: an overwritten word is selected again), redo with the entry's `sel`; when the text changed, the selection is scrolled into view (`revealSel`).
  - `before` is the last editor selection seen since the last commit, noted on `selectionchange`, on every `keydown` (capture) and on `beforeinput` (`app.noteSelection(range?)`; find.js passes the match it replaces). A `MutationObserver` (childList + characterData, not attributes: layout only changes attributes) blocks noting while the DOM differs from the last commit, e.g. between the two halves of a drag-and-drop move.
  - Selections are stored as paths counted the way the HTML re-parses on undo (adjacent text nodes as one, empty text nodes ignored; `pointOf`/`nodeAt`), so they survive text-node splits.
  - Native undo is intercepted: `beforeinput` `historyUndo`/`historyRedo` and Ctrl+Z.
  - Every user-visible change must end with `app.commit(...)`.
- **`.sayfa` files** are the JSON from `app.serialize()`: `app: 'SerbestSayfa'`, the HTML, the images, the used assets, the page settings, `hf` (header/footer) and `styles`. Old files with `pageNumbers: true` load as a centered `{sayfa}` footer.
  - `app.load()` checks the top-level field types before touching the open document, so a broken file can't half-load. User-facing errors are thrown as plain `Error` with a Turkish message; `loadFile` shows those and replaces browser errors (`TypeError`, `SyntaxError`…) with a generic Turkish message.
- **Header/footer** (`state.hf`): `{ header: [left, center, right], footer: [...], firstPage }`. Slots are plain text with `{sayfa}` / `{toplam}` tokens. They are drawn on the pages (`renderPages`), exported as header/footer parts with center/right tab stops and PAGE/NUMPAGES fields (`firstPage` → `titlePg` + empty first-page parts), and imported from the first non-empty paragraph of Word's default header/footer.
- **Autosave** is per tab: IndexedDB key `autosave:<tab id>`, where the tab id lives in `sessionStorage`. Each open tab holds a Web Lock (`serbestsayfa-sekme:<id>`). A new tab adopts only records whose lock is free (closed tabs). The old single `autosave` key is adopted the same way.
  - On `pagehide`, `beforeunload` and `visibilitychange→hidden`, unsaved changes are also written synchronously to `localStorage` (`serbestsayfa-acil:<id>`), because async IndexedDB writes don't survive unload. On startup the newer of the two is used.
  - `beforeunload` asks before closing while there are unsaved changes (cleared by save, open, new, and Word export).
  - Inside the app, New, Open, dropping a file and Close go through async `confirmDiscard()`. With unsaved changes it shows `#saveDialog` (Kaydet / Kaydetme / Vazgeç). Kaydet runs `saveDoc` and continues only if the document really got saved.
  - A save picker can outlast the click's transient activation. `openDoc` then asks for another click instead of calling a file picker Chrome would block.
  - Close is New plus a toast. The empty document's autosave deletes the tab's recovery record.

### Text editing (text.js)

- Formatting uses `execCommand` with `styleWithCSS`. Font size (`fontSize 7` marker trick: apply size 7, then rewrite those marked elements to pt) and sub/superscript (`app.setScript`, which must produce `<sub>/<sup>`, not a `vertical-align` span) turn `styleWithCSS` off for the command.
- **Block structure.** Chrome's list commands can put the list inside the paragraph (`<p><ol>…</ol></p>`). Re-parsing that HTML creates empty paragraphs. `app.normalizeBlocks()` splits such paragraphs after every edit (`ensureContent`); after `load`/undo it also drops the parser's childless `<p></p>`.
- **Line spacing** is stored in CSS `line-height`: a unitless value is a Word "multiple" × the font's single-line factor (`SS.lineFactor`, measured with `line-height: normal`; Calibri 1.221). A value with units is Word's "at least". `app.lineSpacing(block)` / `app.cssLineHeight(block, multiple)` convert. A paragraph without its own line-height takes its style's (`app.styleLineHeight`). Export writes `lineRule="auto"` / `atLeast`.
- **Applying a paragraph style** (`app.setBlock`, Ctrl+Alt+1/2/3) removes direct font, size and colour that cover the whole paragraph (Word's rule), so the paragraph shows its style. It also removes `data-num`, so the heading is numbered again. Bold/italic and paragraph formatting stay. The style dialog (`app.styleDialog`, "Stili değiştir…" in the style box) edits all four styles, "Seçimden al" copies the caret paragraph's formatting.
- **Backspace at a paragraph start** (`backspaceAtStart`) follows Word: first a list item becomes an indented paragraph (the list splits and numbering continues via `ol[start]`) and a numbered heading loses its number (`data-num="0"`), then the first-line indent goes, then the left indent shrinks by 1.27 cm, then a preceding page break (`.pb`) is removed. `outdentBlock` never reduces a numbered heading's numbering indent.
- **Tab** (`tabKey`): multiple paragraphs or a whole paragraph selected → indent/list level up; at a list item's start → level down; elsewhere a tab character. **Shift+Tab**: list item → level up; at a paragraph start → first-line indent, then left indent (`outdentBlock`, shared with Backspace).
- **Paragraph mark.** A selection that ends at offset 0 of the next paragraph (triple-click, Shift+Down) includes the paragraph mark, as in Word. The block after it (B) may be a page break: a triple-click ends right before a `.pb`, and a selection ending at the start of the paragraph right after a `.pb` treats the `.pb` as B (the caret can't be put on it), so the page break is never part of such a selection. `beforeinput` deletions go through `deleteParagraphs()`: whole paragraphs from their start are removed as units (the next paragraph keeps its own style); otherwise the range is trimmed to the end of the paragraph so Chrome doesn't merge the next one in. Enter over whole paragraphs leaves one empty paragraph (`enterOverParagraphs`). Typing, paste and the link dialog call `app.trimParagraphMark()` first; `app.deleteSelection()` deletes with the same rules.
- **Internal clipboard.** Copy/cut from the editor write HTML wrapped in `data-serbestsayfa="paragraf|metin"`. Pasting it skips the sanitizer: whole paragraphs go in as blocks with their own styles (`insertParagraphs`; into a list item, non-list paragraphs split the list with `app.splitTop` and go between: `insertParagraphsInList`), inline copies as a DOM insertion (`insertInline`, keeping the source block's font and size). Don't use `insertHTML` for this: Chrome drops style spans. Anything else is external and goes through `app.sanitizeHTML()`.
- **Removing a list** (list buttons when every selected item is already in that list type) is done by `removeList`/`rebuildList`, not Chrome: each selected item becomes its own paragraph, unselected items are rebuilt at their level and renumbered with `ol[start]`. Creating lists and switching list type are still Chrome's.
- `normalizeBlocks` turns Chrome's `span[style*=vertical-align]` (typing style for sub/superscript) into `<sup>/<sub>` and splits them around nested baseline spans.
- The Paragraph dialog (`app.paragraphDialog`) writes `margin-left/right`, `text-indent` (+ first line, − hanging), `margin-top/bottom` in pt and line spacing. The ruler (`ruler.js`) shows and drags the same indents for the caret's paragraph. For list items both set the list indent (below).
- **List indent** (Word's, measured from the margin): `--li-pad` is the item text's position, `--li-ti` the marker's offset from it (negative = hanging). `css/app.css` gives each nesting depth Word's default (1.27 cm × level, 0.635 cm hanging, up to 9 levels) and makes nested lists start at the margin (a list inside an `li` takes back the item's padding; Chrome's `ol > ol` has no padding), so values never add up.
  - The marker ends with a tab, and tab stops count from the item's text position. `li` has `tab-size: max(1.27cm, hanging)`, so the stop after the marker is the text position (Word: the number's tab goes to the hanging indent, or to the next stop if the number is longer). With the editor's 1.27 cm, a hanging indent over 1.27 cm put a stop between marker and text.
  - A different value is an inline custom property on the list (the whole level) or on the item, in pt rounded to twips (`app.setListIndent`; equal to the default/list value → removed). The ruler and the Paragraph dialog use `app.applyListIndent(items, left, first)`: all items of a list selected → the list, otherwise the items.
  - Tab/Shift+Tab clear the moved items' own values (the new level's indent applies, as in Word); `unlistItem` puts the paragraph at the item's text position; `rebuildList`/`fixList` copy the list's values (`app.copyListIndent`).
  - Chrome copies the item's computed `text-indent` onto the paragraph created by Enter in an empty item. The default value is recognised by `normalizeBlocks`; for others `noteListExit`/`dropListIndent` note the item's value before Enter and remove only that copy (a value match alone could strip a real hanging indent).
  - Import (`docximport.js` `listInd`): paragraph `w:ind` first, then the numbering level's (numbering on the paragraph) or the style's (numbering from the style); the list takes its first item's value, other items their own; an undefined level with no indent anywhere keeps the default. Word paste: `margin-left`/`text-indent` of `mso-list` paragraphs. Export: `w:ind` on the paragraph when it differs from the level definition (720 × (level + 1) / 360).
- Paste goes through `app.sanitizeHTML()`. It reduces Word and web HTML to `p/h1-3/ul/ol/li/b/i/u/s/sub/sup/span/a[href]/br` and `.pb`, turns Word's fake list paragraphs (with their `mso-list` levels) into nested lists, and keeps Word's pt/cm paragraph indents and spacing and `line-height: %`.
  - Word's `tab-stops:` become `data-tabs`, and `mso-tab-count` spans become tab characters.
  - The fake numbers of Word's numbered headings (`mso-list:Ignore`) are dropped.
- Keep editor HTML within that vocabulary, plus `ol[start]`, `ol/ul[data-lf]`, the `--li-pad`/`--li-ti` inline properties on `ol/ul/li`, `[data-tabs]`, `span.tab` and `h1–h3[data-num="0"]`; the internal clipboard keeps those. Export and import only understand it. Word's pasted lists get `data-lf` from the fake marker (`wordListFormat`: "a)", "IV.", "(1)", "1.1.", Symbol/Wingdings bullets).
- **Symbols:** the Ω button opens `#symbolPanel` (recent symbols in localStorage first, then the defaults; grouped symbols; a hex code field); `app.insertSymbol(c)` inserts with `insertText`. Alt+X converts the hex code before the caret (after `U+`, a valid 5–6 digit non-BMP code, else the last 4 digits) to the character, or the character to its code (`codeSwap`).
  - Target (`symbolTarget`): a caption being edited (`.cap-text`), a text input in a dialog or panel (header/footer, find/replace; not the toolbar boxes), else the editor. Panel buttons don't take focus (`mousedown` → `preventDefault`), so a caption keeps editing. From the code field the last text field and its caret are used (captured in a capture-phase `blur`, before the caption's own blur ends editing); a caption is reopened with `app.editCaption`.
  - In the editor both are separate undo steps (`command()` + `commit('edit')`); in fields the field's own undo applies.
  - Modal dialogs make everything outside inert: `app.symbolPanel(host)` moves the panel into the dialog (Page Setup's "Ω Simge…" button, or automatically when the panel was open) and back on `close`. `app.placeSymbolPanel()` stacks it under the find panel when both are open.
- **Links** are `<a href>` with `http:`, `https:` or `mailto:` only (`app.safeHref`; anything else stays plain text on paste, `.sayfa` load and import). Ctrl+K / the toolbar button open `app.linkDialog()` (add, edit, remove; `www.…` → `https://`, e-mail → `mailto:`); Ctrl+click opens a link. Export: `w:hyperlink r:id` + external relationship + `Hyperlink` character style (blue/underline come from the style, not direct formatting). Import: `w:hyperlink`, `w:fldSimple` and complex `HYPERLINK` fields; the Hyperlink style's look is not copied onto link runs.
- **History batching.** The `input` handler commits every edit. `batching` and `dragMove` (a drag-and-drop move: `deleteByDrag` then `insertFromDrop`, or `dragend` if dropped outside) suppress those commits; the operation commits once at the end.
  - App commands run their `execCommand`s through `command(fn)` (batching + `withoutBands`), so Chrome's `input` events don't write steps and the command's final `afterFormat()`/`commit()` records one (e.g. Shift+Tab's outdent plus the span cleanup, or paste over a selection: delete + insert). `afterFormat` doesn't commit while batching (nested commands). The format painter sets `batching` itself.
- **Format painter** (`app.formatPainter`, `app.copyFormat`/`app.pasteFormat` = Ctrl+Shift+C/V): copies the character format at the selection start and, for a caret/paragraph-mark/multi-paragraph selection, the paragraph style and inline paragraph properties. Applying: paragraph style first, then `removeFormat` + only the character properties that differ, then paragraph properties last (Chrome's `removeFormat` also clears `text-align`).
- **Auto-link** (`autoLink`): after a space, tab or Enter, a preceding web/e-mail address becomes `<a href>` as a separate history step, so Ctrl+Z removes only the link.
- **Font boxes** are editable combos (`.combo`: the list is the `<select>` underneath, the value is typed in the `<input>` on top; Enter applies, Esc/blur reverts). Sizes are rounded to half points (1–1638). `FONTS` in `text.js` is the curated font list: the dropdown shows the installed ones (canvas width check), paste keeps all of them.
- Find & replace (`find.js`) highlights matches with the CSS Custom Highlight API (no DOM changes) and edits text nodes directly. Replace All also covers captions and is one undo step.

### Objects (objects.js)

- Pointer handling runs in the capture phase on `#doc`. The `mousedown` that follows is cancelled through `blockMouse`, so text focus and selection aren't disturbed.
- Images behind the text can be selected only where no glyph is under the pointer (`isOverText`), or with Alt held.
- Captions are edited in place with `contenteditable="plaintext-only"`. Global key and paste handlers skip `.cap-text`. Enter ends the edit and stops propagation: the image stays selected, and the document's Enter/F2 handler (edit the selected image's caption) would reopen it.
- **Inserting images** (`makeImage`, `caretSpot`): at a paragraph start the image goes on the caret line (the paragraph continues below it); otherwise below the caret's line (`caretLineRect` + `WRAP_DIST` + 2 px, because Chrome can round the line box up). Images inserted together stack below each other; one that doesn't fit goes to the top of the next page.
- **Crop mode** (`app.toggleCrop`, "Kırp"): crop handles cut the frame while the image stays put (`fullImage`/`applyCrop`, in the image's own rotated axes); dragging the image pans it under the frame. Esc/Enter end it; the overlay shows the cut-away part faded (`.crop-ghost`, `clip-path` with a hole).
- After Esc, Delete or the delete button, focus returns to the text at the last caret (`returnToText`). Cut doesn't, so paste keeps its target rule (`pasteTarget`).

### Shell (main.js)

- Dialogs are `<form method="dialog">`. Enter submits with the **first submit button**, so "Vazgeç" buttons are `type="button" value="cancel"` and closed by a shared handler; keep "Uygula" the first submit button.
- **Word count.** The status bar shows the total (`innerText`, debounced after changes). While text is selected it shows "selected / total" (`Selection.toString()`, debounced on `selectionchange`). Clicking it or Ctrl+Shift+G opens `app.wordCountDialog()`: pages, words, characters, paragraphs and lines of the selection or the document.
- Header/footer is edited in the page setup dialog: `pageSetup(focus)` opens it on a slot (`h0…h2`, `f0…f2`). Double-clicking a page's top/bottom margin opens the slot under the pointer (`hfSlotAt`); the toolbar has a header/footer button.

### DOCX export (docx.js)

- The ZIP is store-only (`zip.js`).
- Images are `wp:anchor` with `relativeFrom="page"`. Word and LibreOffice put an anchored object on the page of its anchor *paragraph*, so a page's images are anchored at the start of a paragraph that **starts** on that page, the one nearest the page's vertical middle (`placeAnchors`), which tolerates a few lines of line-breaking drift. If no paragraph starts on the page, the page's first character is used.
  - Paragraph starts and `findPageStarts` are measured from the live layout, so **export depends on the current DOM layout** (it calls `app.paginateNow()` first).
  - A page with no text gets a holder paragraph with `pageBreakBefore`, and its images are written with `wrapNone`.
- Lists use Word's default indents: text at `720·(level+1)` twips, marker 360 twips to its left. A nested list written directly inside a list (Chrome's indent form `<ul><li/><ul>…</ul></ul>`) is exported as the next level.
  - Lists with a format (`data-lf`, their own or an ancestor's) get their own `abstractNum` (id 3+, shared by equal definitions). The list's levels are shifted to Word levels. Bullets are written with Word's Symbol/Wingdings characters (`BULLET_FONT`). Each list has its own `numId` with a `startOverride` on its level.
  - Lists without any format are written exactly as before (abstractNum 0/1).
- **Heading numbering** is written as `abstractNum` 2 (`multilevel`, levels linked with `w:pStyle` to Heading1–3) plus `w:numPr` in those styles.
  - Unnumbered headings and the continuation of a split heading get `w:numId 0`.
  - A numbered heading gets `w:ind` only when it has its own inline indent.
- `data-tabs` becomes `w:tabs` in `pPr`, after `numPr` and before `spacing` (px × 15 = twips).
- Captions are `wps` text boxes inside `mc:AlternateContent`, with the `Caption` style and a `SEQ` field; chapter-numbered captions add `STYLEREF 1 \s` + the separator and `SEQ … \s 1`.
- Page breaks become `pageBreakBefore` on the next paragraph.
- Run and paragraph formatting is read from computed styles, but only what differs from the paragraph's style (`styleOf`: Normal or Heading 1–3) is written as direct formatting. styles.xml is generated from `state.styles` (Normal in docDefaults + Normal style, headings basedOn Normal), so changing a style in Word changes the text.
- Keep OOXML child-element order in `pPr`, `rPr` and `wp:anchor`. Word rejects files with the wrong order.

### DOCX import (docximport.js)

- **Reading:** unzips with `DecompressionStream('deflate-raw')` and resolves styles, theme fonts and numbering. `parseNumbering` simulates Word's counters; a skipped higher level shows 0, as in Word ("1.0.1").
- **Heading numbering** (`headingNumbering`):
  - The list is the one linked to the Heading 1–3 styles, or else the one most headings use.
  - Heading k on level k−1 of that list is left to the editor's counters (`styles.num`; the start value comes from the first numbered heading). Other headings of a numbered level get `data-num="0"`.
  - The counters are replayed against Word's numbers. If any differs (a restart in the middle) or a format isn't supported, the numbers are written into the text as before.
- **Tab stops:** the `w:tabs` of the style chain and the paragraph are merged into `data-tabs` (`clear` removes a stop). Bar and num stops are dropped.
- **HTML with placeholders:** drawings become `<span class="ph">` placeholders in the generated HTML.
- **Styles:** Word's Normal (docDefaults + default paragraph style) and "heading 1–3" styles become `state.styles` (`importStyles`; a missing heading style keeps the built-in one). `TAG` (the defaults paragraphs and runs are compared with) is derived from them per import.
- **Paragraph formatting:** where a paragraph's own Word style differs from those, its font and size go on the block element itself; runs only get spans where they differ (font, size, colour, bold and italic are compared with the paragraph's style). `auto` line spacing becomes a multiple × `SS.lineFactor` (unspecified = single). List level = rank of the effective left indent within a run of list paragraphs (so Word's single-level "List Bullet 2" becomes level 2); an ordered list that resumes after other paragraphs gets `ol[start]`.
  - A new top-level list starts when the Word list definition (abstractNum) changes or the numbering restarts, so separate Word lists stay separate.
  - Each level's `numFmt`/`lvlText` (Symbol/Wingdings bullets mapped to Unicode by `bulletChar`) becomes the owner list's `data-lf`, with placeholders shifted to the editor level; default levels are left out.
- **Placement:** after `app.load()`, `placeAll()` goes through the placeholders in document order. For each one it finishes pagination up to the placeholder's page, measures the placeholder, creates the image object, and re-runs layout before the next.
  - Inline images become `topbottom` at their line.
  - Anchored images follow Word's positioning rules.
- **Caption numbers:** a two-part number ("Şekil 2.1", `.` `-` `–` `—`, from fields or typed) makes the caption chapter-numbered. After placement the editor's numbers are compared with the document's; multi-part numbers that differ are counted in the import summary (`capRenum`). Plain numbers aren't checked, because Word's cached SEQ results are often stale.
- **Caption attachment** looks for:
  - A paragraph right after (or right before) an image-only paragraph that matches `CAPTION_RE` or uses the `caption` style.
  - A caption-like text box in the same drawing or group.
- The document name is the file name (not `dc:title`). Only "heading 1–9" styles and explicit outline levels become headings; "Title" is a Normal paragraph with its formatting.
- **Lossy conversions:** tables become tab-separated paragraphs. EMF/WMF/TIFF images, charts, footnotes and header/footer content beyond the first text line are skipped, and `SS.importSummary()` tells the user how many.
