# Changelog

## Unreleased

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
