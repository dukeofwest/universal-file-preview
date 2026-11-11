# Changelog

## 0.4.0 - 2026-11-11

This release adds OpenDocument text previews and fixes two resource-exhaustion issues in the handling of untrusted files. It is a minor version because the public `FileFormat` type gains three members and the default renderer registry serves them; existing code that does not switch exhaustively over `FileFormat` works without changes. Upgrading is recommended for applications that preview files from untrusted sources. The package still has no runtime dependencies.

### New features

- **OpenDocument text previews**, rendered through the same viewer as DOCX. Paragraphs and headings keep their outline level, text and paragraph styles are resolved through their inheritance chains, and spans, spaces, tabs and line breaks are read. Lists carry their numbering, including nested levels, restarts and continuations; tables carry their spans, column widths and cell styles. Images stored in the package, or inline in a flat document, are shown at their declared size with their alternative text. Page size, margins, headers and footers come from the document's master page, including a distinct first page.
- A template is labelled as a template and shown as saved, with its placeholders never filled in.
- Content that cannot be shown faithfully is reported in the preview instead of disappearing (see Worker and diagnostics improvements).

### New OpenDocument format support

| Format | Extension | Media type |
| --- | --- | --- |
| Text document | `.odt` | `application/vnd.oasis.opendocument.text` |
| Text template | `.ott` | `application/vnd.oasis.opendocument.text-template` |
| Flat text document | `.fodt` | single-file XML, no package |

`detectFormat` returns these for text documents identified by the package's `mimetype` entry, by flat-XML content, by MIME type or by extension. A document is identified by its own declared media type rather than by its file name.

This is a defined subset of the OpenDocument text specification, not complete specification support. See the known limitations below and [docs/format-support.md](docs/format-support.md) for what is read.

### Worker and diagnostics improvements

- Worker mode covers the new formats, so a text document can be parsed off the main thread like any other. The worker's format list and the model kind it expects for each format are now a single explicit table, so a format cannot be added to the protocol without stating which model it returns.
- Two diagnostics are new: a document that is a template is reported as one, and content the preview cannot show faithfully is counted and named above the document instead of being dropped silently. This covers notes, comments, text boxes, embedded objects, charts and tracked changes.

### Security fixes

- **OpenDocument spreadsheets: XML element amplification.** The number of XML elements a document's content may hold is now bounded by the workbook's cell limit (250,000 plus eight per permitted cell, so 4,250,000 at the default limit) instead of by the size the part is allowed to have, which stood at about 33.5 million elements whatever a file contained. A 115 KiB package of many empty elements could keep the parser busy for about 11 seconds of CPU time; it is now refused while being read. **Affected versions: `>= 0.2.0, <= 0.3.0`**, every version with OpenDocument spreadsheet support.
- **Namespace resolution: unbounded copying.** Entering an element that declares a namespace copied every binding then in scope, and nothing bounded that work. The nesting and attribute limits together allow more than 131,000 bindings at once, so a few hundred elements below such a stack cost tens of seconds and more than a gigabyte of memory while staying inside every element and token limit. Entering an element now costs only the declarations it makes, and a prefix is resolved by walking the enclosing scopes and remembering the answer at each one. A 560 KiB spreadsheet of 433 elements went from about 17 seconds and 1.1 GiB to 0.19 seconds and 14 MiB. **Affected versions: OpenDocument spreadsheets `>= 0.2.0, <= 0.3.0`; SVG images, standalone and embedded in XLSX and PPTX, `= 0.3.0`** only, the version whose SVG sanitizer parses with the library's own XML layer. Word documents do not display SVG images, so DOCX is not affected on that path.

A compression-ratio limit is deliberately **not** applied to OpenDocument parts. Ordinary spreadsheets compress far beyond the 100:1 ratio used for Office Open XML parts (186:1 to 218:1 measured for a legitimate 500,000-cell workbook), and an attack can be padded below any threshold that keeps such files working, so the ratio separates neither case. The element budget is the control, and it is independent of how the file is compressed: the test suite refuses the same attack compressed, padded below 100:1, and stored uncompressed.

### Compatibility changes

- The public `FileFormat` type now also includes `"odt"`, `"ott"` and `"fodt"`. TypeScript code that switches exhaustively over `FileFormat` needs a branch for them.
- The default renderer registry now serves those formats, so registering another renderer for them in `createDefaultRendererRegistry()` throws a duplicate-format error; use `new RendererRegistry()` for a fully custom setup.
- A file whose name or MIME type says it is an OpenDocument text document but whose content is not one now resolves as that format and is refused by the parser, rather than resolving as `"unknown"`.
- The element budget can refuse OpenDocument spreadsheets that 0.3.0 accepted. A workbook may now hold on average eight XML elements per permitted cell. A plain 500,000-cell workbook holds about two elements per cell, and the same workbook with an annotation or several spans in every cell about four to five, so ordinary documents are unaffected; a document that is both at the cell limit and far denser than that is refused with a `resource-limit` error. Every element in the content stream counts, including ones the viewer discards, such as styled empty cells outside the drawn grid, content-free covered cells, and skipped pivot-table or conditional-format subtrees, so a file consisting mostly of those can now be refused. A spreadsheet that writes each empty cell individually instead of using a repeat count is the most likely to reach the limit without holding much content.

### Known limitations

- Master documents (`.odm`), master templates (`.otm`) and web templates (`.oth`) are not read and stay unrecognised. OpenDocument presentations (`.odp`, `.otp`, `.fodp`) and drawings (`.odg`) are not supported either.
- Text documents have no pagination, exactly as for DOCX: the preview is a continuous flow, and headers and footers are shown once rather than per page.
- Notes, comments, text boxes, embedded objects, charts and tracked changes are reported rather than rendered.
- Linked sections, embedded objects and template references are never fetched, and an image reference with a scheme, an absolute path or a parent-directory step is not followed.
- A list numbered through its paragraph style's `style:list-style-name`, rather than through the list's own `text:style-name`, is shown with bullets instead of its numbering.
- A table is limited to 10,000 rows, 256 columns and 1,000 paragraphs in a cell. Content past those limits is left out without a notice, unlike the content kinds listed above.
- A text document parsed on the main thread cannot be interrupted; worker mode can be, as for the other formats.
- The resource limits are fixed and cannot yet be changed by the application.
- Parsing runs on the page's main thread unless worker mode is enabled. See [docs/security.md](docs/security.md) for the remaining gaps; this package is not claimed to be free of vulnerabilities.

## 0.3.0 - 2026-11-07

This release adds an optional worker mode, limits the size of every parsed document, and lets the viewer run under a strict Content Security Policy and Trusted Types. It is a minor version because the public API gains options and an entry point; existing code works without changes. Some limits are stricter than in 0.2.x and can refuse files that 0.2.x accepted (see Compatibility). Upgrading is recommended for applications that preview files from untrusted sources.

### Added

- **Worker mode:** the new `worker` option parses DOCX, XLSX, PPTX, ODS, OTS, and FODS files in a Web Worker that the application creates from the new `./worker` entry point (`dist/worker.js`). The page stays responsive while a document is parsed; a parse that exceeds the timeout (30 seconds by default) is stopped and fails with `resource-limit`; replacing or destroying the viewer stops a running parse. `fallback: "main-thread"` parses on the main thread when the worker cannot be started. Results from the worker are checked before rendering. Without the option, parsing runs on the main thread as before. The installation guide describes the setup for bundlers and without one.
- `load()` accepts an optional second argument with an `AbortSignal` (`load(file, { signal })`); a cancelled load shows "Preview loading was cancelled." instead of staying at "Loading preview...".
- The `styleNonce` option sets the nonce of the viewer's `<style>` element in browsers without constructable stylesheets.

### Security

- The parsed document model of every format (DOCX, PPTX, XLSX, ODS, OTS, FODS) is now limited while it is built: at most 6,000,000 objects and arrays, 67,108,864 UTF-16 code units of text, 256 MiB of image data, and a nesting depth of 64. A document that exceeds a limit is refused with a `resource-limit` error instead of being shown.
- The viewer stylesheet no longer needs `'unsafe-inline'` in a Content Security Policy: in browsers with constructable stylesheets (tested in Edge, Chrome and Firefox) it is adopted by the container's document or shadow root through `adoptedStyleSheets`, shared by all viewers there and removed with the last one. Other browsers get a `<style>` element in the container as before.
- The viewer uses no markup sink: its icons are built with DOM methods instead of `innerHTML`, and SVG images are sanitized without the browser's `DOMParser` (parsed by the library's own XML parser, reduced to the same allowlist as before and written out with escaping). It therefore works under Trusted Types (`require-trusted-types-for 'script'`) without a policy, and sanitizing malicious SVG no longer causes Content Security Policy reports.
- Fixed a case in which a crafted XLSX table could make the preview stop responding, and a case in which unusual DOCX list levels made the preview fail. Both were found by the new fuzzing tests.
- The download button always saves the original file as `application/octet-stream`, whatever type the source declared, and its `blob:` URL is revoked shortly after the download starts. Error messages no longer quote ZIP entry or package part names from the file (the name is kept in the error's `cause`), so a file cannot put its own wording into the viewer's status line. Relationship references in Office files resolve their namespace prefix, a prefixed `TargetMode="External"` is recognised, and lookups keyed by document text use only own properties.

### Compatibility

- The new limits can refuse some files that 0.2.x accepted:
  - Spreadsheets (XLSX and OpenDocument) may have at most 500,000 cells (previously 1,000,000). A workbook with more cells is refused with a `resource-limit` error.
  - XLSX text is counted more strictly: a cell's formatted display text now counts in addition to its value, as do the texts of hyperlinks, comments and other sheet content. A workbook in which long texts are repeated in many cells can therefore be refused although 0.2.2 accepted it (for example, a 32,767-character shared string may now be used by at most 1,023 cells, where 0.2.2 allowed 2,048).
- SVG images with a DOCTYPE (with or without entity declarations) are no longer shown; 0.2.x expanded internal entities and removed the DOCTYPE. SVG containing characters XML does not allow is not shown either.
- Format detection checks the signatures at the start of a file first: a `%PDF-` marker later in the first KiB now counts only when no other format matched (0.2.x treated, for example, a ZIP package whose first entry name contained `%PDF-` as a PDF).
- In browsers with constructable stylesheets the viewer no longer adds a `<style>` element to the container, and its stylesheet now comes after all of the page's own stylesheets in the cascade. Application rules that override `ufp-*` classes from a stylesheet placed after the viewer need a more specific selector.

## 0.2.2 - 2026-11-06

This patch release fixes a resource-exhaustion issue in XLSX previews that 0.2.1 did not cover. Upgrading is recommended for applications that preview files from untrusted sources. There are no API changes.

- **Security:** XLSX text is now limited: one cell may hold at most 32,767 characters (Excel's limit), and the text of all cells together, counted for every cell that uses it, at most 67,108,864 UTF-16 code units. A workbook that exceeds either limit is refused with a `resource-limit` error instead of being shown.
- Formula text results longer than 32,767 characters are `#VALUE!`, as in Excel, and `SUBSTITUTE` with an empty search text returns the text unchanged, as in Excel.
- Number formats that would produce more than 32,767 characters of text are refused.

## 0.2.1 - 2026-11-06

This patch release hardens the handling of untrusted documents. Upgrading is recommended for applications that preview files from untrusted sources. There are no API changes; the new limits below can refuse or simplify unusual files.

- **Security:** hardened the handling of embedded images and spreadsheet styles, and added resource limits for DOCX, PPTX, and XLSX files (total decompression, XLSX calculation, and image size). Upgrading is recommended for applications that preview files from untrusted sources.
- Embedded images are shown only when their content matches a supported image type; media that does not match is no longer displayed. Unusual font names in documents may now fall back to a generic font.
- The viewer root now uses CSS containment (`contain: layout paint style`, `isolation: isolate`), so elements with `position: fixed` placed inside `.ufp-root` are positioned and clipped relative to the viewer.
- Behaviour changes from the new limits: ZIP files whose records overlap or disagree between local and central headers are refused; Office XML parts that expand more than 100 times their compressed size are refused; a workbook with more than 1,024 sheets or 1,000,000 cells is refused; a slide or sheet listed more than once is shown once; images larger than 40 megapixels are not shown; and formulas beyond the calculation limit are shown as not calculated.
- Spreadsheet views no longer keep image data for drawings that have scrolled out of view and been redrawn.
- The security documentation now lists the remaining known gaps (synchronous parsing, fixed limits) and recommends a deployment and Content Security Policy for untrusted files.

## 0.2.0 - 2026-11-04

This release adds OpenDocument spreadsheet previews (ODS, OTS, FODS). It is a minor version because the public `FileFormat` type gains three members; see the compatibility notes below.

- **Compatibility note:** the default renderer registry now includes a renderer for `"ods"`, `"ots"`, and `"fods"`, so registering another renderer for those formats in `createDefaultRendererRegistry()` throws a duplicate-format error; use `new RendererRegistry()` for a fully custom setup.
- **Compatibility note:** the public `FileFormat` type now also includes `"ods"`, `"ots"`, and `"fods"`. `detectFormat` returns them for OpenDocument spreadsheets (by package `mimetype` entry, flat-XML content, MIME type, or extension). TypeScript code that switches exhaustively over `FileFormat` (for example with a `never` check) needs cases for the new members.
- Added OpenDocument spreadsheet previews (`.ods` and `.ots` templates) in the existing workbook grid. Cells show the text the producing application saved, so values are never reformatted for the viewer locale; formula results are shown as last saved (never recalculated) with a notice, and hidden or truncated content is reported. The demo includes a sample ODS file.
- ODS number formats (data styles) are now read: negative amounts and other conditional sections get their colour (for example red), values without saved text are formatted in the document's own locale, and saved text is checked against the format, with a notice when they differ (the saved text is still what is shown). Formats that cannot be reproduced exactly, such as embedded text, fill characters, non-Gregorian calendars, General, optional decimals and fixed fraction denominators, are left to the saved text.
- ODS previews open with the view saved by LibreOffice (`settings.xml`): the active sheet, frozen rows and columns, zoom, and whether grid lines and row/column headers are shown. These settings are optional; if they cannot be read, the defaults are used.
- Added previews of flat OpenDocument spreadsheets (`.fods`, a single XML file), with the same results as the equivalent `.ods` package. The demo includes a sample `.fods` file.
- SVG previews keep fills and strokes that reference a gradient inside the same image (`fill="url(#id)"`); previously the sanitizer removed them and such shapes rendered black. Remote, `data:` and `javascript:` references are still removed.
- `detectFormat` no longer returns JavaScript object properties for file names or MIME types such as `x.constructor`; they are reported as `"unknown"`.
- Added published-package integration applications for Angular, React, and Vue, plus a unified GitHub Pages site and deployment workflow.
- Documented framework lifecycle patterns, local Pages builds, and the required repository Pages setting.
- Sharpened npm metadata to describe implemented browser previews and removed unverified framework/PDF-viewer keywords.
- Updated package, security, and release documentation to reflect the published `0.1.1` package and current repository status.
- Added a technical content and organic adoption plan with examples based on the public API.
- Made the local demo server's host and port configurable and documented hosting setup without enabling an automatic deployment.

## 0.1.1 - 2026-10-29

- Published the scoped package as `@dukeofwest/universal-file-preview` with corrected public package metadata and ESM distribution entries.

## 0.1.0 - 2026-10-29

- Published publicly to npm as `@dukeofwest/universal-file-preview`.
- Expanded release documentation and aligned format-support claims with current parser coverage and automated test evidence.
- Rebuilt XLSX rendering for Excel fidelity: cell fonts, theme/indexed/tinted colors, solid, pattern and gradient fills, all border styles, alignment, wrapping, indent, rotation, shrink-to-fit, rich text, merged cells, and Excel-exact column widths and row heights.
- Added an Excel number-format engine (sections, conditions, colors, scaling, fractions, scientific, dates/times, elapsed time, accounting fill) that uses the viewer locale for separators, month/day names, and built-in short dates.
- XLSX conditional formatting (expression rules with defined names, cell/text/blank/error/duplicate/top/average/date rules, color scales, data bars, icon sets), built-in and custom table styles with filter buttons, frozen panes, hidden rows/columns/sheets, saved scroll position and zoom, notes, web hyperlinks, pictures, and shapes.
- XLSX drawings render shapes and shape groups through the DrawingML element renderer shared with PPTX: preset (including teardrop) and custom geometry, gradients, rotation, 3D camera tilt with perspective and Office-like lighting, WordArt arch text, reflections, and text shadows.
- XLSX wrapped text uses Excel's per-font line pitch (also used for auto-fitted rows), a trailing line break keeps its empty line, synthesized bold is widened like GDI, and table filter columns with hidden buttons show none.
- Franklin Gothic Book (an Office cloud font browsers cannot load) falls back to a local face matched to its widths and line spacing, with Excel's 9 px column digit width.
- XLSX slicers render as in Excel (caption, buttons in columns, selection from the table filter or pivot cache, custom and built-in slicer styles) instead of the "not supported" fallback text; sparklines (line, column, win/loss) are drawn in their cells; data validation input messages appear next to the selected cell.
- XLSX pictures in TIFF format are converted to PNG (uncompressed, PackBits, LZW and Deflate TIFFs), keeping their embedded color profile.
- Gill Sans MT (an Office cloud font) falls back to Corbel with lining figures, size-matched to its widths, with Excel's hinted digit width, line pitch and indent step; row heights are truncated to whole pixels like Excel; arrow shapes center their text in the arrow shaft; cell fonts that match the Normal style let table-style font colors show through.
- XLSX text overflows into empty neighbouring cells, numbers that do not fit show `####`, and General numbers shorten like Excel.
- Uncached XLSX formula cells are computed with a small built-in evaluator; shared strings ignore phonetic runs and decode `_xHHHH_` escapes.
- Font fallbacks prefer metric-compatible faces (Calibri for Aptos Narrow, Carlito, Caladea, Liberation) when Office fonts are not installed.
- Improved PPTX fidelity: shapes and lines now render as SVG with true preset geometry (arrows, chevrons, hexagons, rounded/snipped rectangles, stars, connectors), adjust values, flips, outlines, dash styles, arrowheads and slide-space shadows.
- PPTX shapes honour theme style references (fill, line, effect and font colour), `useBgFill` background-filled shapes, picture geometry clipping, and master-slide shapes, backgrounds (including `bgRef`) and placeholder positions.
- Corrected PPTX colour transforms (tint/shade in linear RGB, luminance/saturation/hue in HSL), gradient direction, radial gradients, preset/scRGB/HSL colours, and `normAutofit` font scaling.
- PPTX text: SmartArt text rectangles stay upright, first-paragraph spacing and line spacing follow PowerPoint, empty paragraphs hide bullets, bullets use hanging indents, and old-style or light fonts fall back to similar local faces.
- Improved DOCX style fidelity for theme-aware fonts, font fallbacks, paragraph shading and borders, line spacing, underline styles, character spacing, and anchored DrawingML shape groups.
- Added DOCX horizontal table banding from custom table styles and preserved whole words in table cells when columns are tight.
- Honored explicit automatic cell shading so table band colors remain confined to cells without direct shading.

- Initial browser package, native image renderer, file detection, binary reader, and typed lifecycle API.
- Added original ZIP/DEFLATE and XML foundations and basic DOCX, XLSX, and PPTX subsets.
- Documented format status, specifications, security limits, and known gaps.
