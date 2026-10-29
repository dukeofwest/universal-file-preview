# Changelog

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
