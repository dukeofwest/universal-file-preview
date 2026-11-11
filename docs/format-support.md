# Format support and implementation matrix

This document records the current implementation rather than promising complete compatibility. Parsers treat input as untrusted and return an error when the implemented subset cannot be read safely.

| Extension(s) | Detection | Parsing | Rendering | Implemented subset and material gaps | Automated tests / status |
|---|---|---|---|---|---|
| `.png`, `.jpg`, `.jpeg`, `.webp`, `.gif` | Signature, MIME, or extension | Browser handles image decoding; no project image codec | Yes, via browser `<img>` decoding | GIF animation is browser-provided. Codec support depends on the browser. | Detection has synthetic tests; browser decoding and viewer presentation are not automated in this suite. Partial. |
| `.svg` | XML root, MIME, or extension | Project XML parser plus allowlist sanitizer (no DOCTYPE) | Yes, via browser `<img>` decoding | Only allowlisted elements in the SVG namespace and allowlisted attributes with strict value rules survive; the result is parsed and validated again. Scripts, event handlers, style attributes, external references, and unsupported SVG features are removed. The same sanitizer handles SVG embedded in DOCX, PPTX, and XLSX. | Detection is tested; `scripts/browser-security.mjs` runs a sanitizer table (namespace confusion, escaped `url()` values, entities, legitimate drawings), compares the sanitizer with the earlier DOM-based one (same trees and pixels), and checks that sanitized SVG runs no script when opened as a document; `tests/svg-sanitizer.test.mjs` covers escaping and invalid characters. Partial. |
| `.docx` | ZIP signature plus extension/MIME | Yes, bounded OOXML package and Word subset | Yes, DOM/SVG renderer | Paragraph/run styles, numbering, tables, page dimensions/margins, headers/footers, selected images, and basic anchored DrawingML. No automatic pagination; incomplete floating layout, fields, comments, footnotes, tracked changes, advanced table-style inheritance, and full Word typography. | Four synthetic parser fixtures; no automated browser rendering or visual-regression tests. Experimental subset. |
| `.xlsx` | ZIP signature plus extension/MIME | Yes, workbook and worksheet subset | Yes, styled, absolutely positioned worksheet grid | Cell styles, number formats, rich text, dimensions, merges, frozen panes, conditional formatting, tables, slicers, sparklines, notes, web links, selected drawings, TIFF conversion, and limited formula evaluation. Charts, many formulas, right-to-left sheets, outline controls, and some effects/styles are unsupported or approximated. | Eight synthetic parser/algorithm fixtures across `xlsx.test.mjs` and `ooxml.test.mjs`; no automated browser rendering or visual-regression tests. Experimental subset. |
| `.pptx` | ZIP signature plus extension/MIME | Yes, presentation, slide, theme, layout, master, and selected drawing subset | Yes, slide renderer with keyboard and previous/next navigation | Text, backgrounds, theme references, placeholders, many shapes/lines, images, tables, selected SmartArt drawing content, gradients, and some effects. No full PowerPoint fidelity; charts, animation, transitions, and embedded media are not implemented. | Two synthetic parser fixtures; no automated browser rendering or real-file visual regression. Experimental subset. |
| `.ods`, `.ots` | ZIP signature plus the package's `mimetype` entry, MIME, or extension | Yes, OpenDocument spreadsheet subset: `content.xml` is streamed; `styles.xml`, `meta.xml`, and `settings.xml` are read | Yes, the same worksheet grid as XLSX | Typed cell values (number, percentage, currency with its ISO code, date, time, boolean, string, error), the producer's saved display text, cell styles with inheritance, column widths and row heights, merges, repeated rows and columns, hidden and filtered rows, columns, and sheets, rich text runs, web and `mailto:` links, comments, named ranges, number styles (negative colours, and formatting in the document locale where no text was saved), and view settings (active sheet, frozen panes, zoom, grid lines, headers). Formulas are not recalculated: saved results are shown, and cells without one show "Cannot calculate". Charts, images and shapes, conditional formatting, pivot-table behaviour, external data, data validation, outline controls, right-to-left sheets, and encrypted documents are not supported. Unsupported or approximated content is listed in a notice above the grid. OpenDocument presentations (ODP) are not supported. | Node tests across `ods.test.mjs`, `ods-financial.test.mjs`, `ods-number-styles.test.mjs`, `ods-settings.test.mjs`, `odf-core.test.mjs`, `odf-package.test.mjs`, `detection-odf.test.mjs`, and `ods-renderer.test.mjs`, all on synthetic LibreOffice-shaped fixtures. Demo-sample rendering is compared by the manual browser regression script. No files from real producers are included yet. Experimental subset. |
| `.fods` | Flat-XML sniff (office namespace and spreadsheet media type near the start) with the `.fods` extension or the spreadsheet MIME type | Yes, the whole file is streamed once | Yes, the same worksheet grid | As for `.ods`. Embedded images and objects are not kept. | `fods.test.mjs` checks that flat files give the same workbook as the equivalent `.ods` package across value, layout, style, formula, text, unsupported-feature, and view-setting fixtures. Experimental subset. |
| `.odt`, `.ott` | ZIP signature plus the package's `mimetype` entry, MIME, or extension | Yes, OpenDocument text subset: `content.xml` is streamed a paragraph at a time; `styles.xml` is read as a tree | Yes, through the same renderer as DOCX | Paragraphs, headings with their outline level, text and paragraph styles (including inherited ones), spans, spaces, tabs and line breaks, lists with their numbering, tables with spans and styles, images stored in the package, page size and margins, headers and footers. A template is labelled and shown as saved. | Master documents (`.odm`) and web templates (`.oth`) are not read. Notes, comments, text boxes, embedded objects, charts and tracked changes are reported rather than shown. There is no pagination, as for DOCX. |
| `.fodt` | Flat-XML sniff (office namespace and the text media type as a whole attribute value) with the `.fodt` extension or the text MIME type | Yes, the whole file is streamed once | Yes, the same renderer as `.odt` | The same as `.odt`, with images read from the inline base64 data. | The same as `.odt`. |
| `.pdf` | Signature, MIME, or extension | No | No built-in renderer | Detection only. | Detection covered by a synthetic test. Unsupported for preview without a custom renderer. |
| `.doc`, `.xls`, `.ppt` | Extension/MIME; a CFB signature is only used with subtype metadata | No CFB or legacy format parser | No built-in renderer | Detection only; these are not equivalent to their OOXML counterparts. | Detection covered by a synthetic test. Unsupported for preview without a custom renderer. |

Detection does not imply the package can render a file. The default renderer registry contains handlers only for the image formats, DOCX, XLSX, PPTX, OpenDocument spreadsheet (ODS, OTS, FODS) and OpenDocument text (ODT, OTT, FODT) rows marked as implemented above. A detected PDF or legacy Office file produces an `unsupported-format` error unless the application registers a custom renderer.

## Format structures and planned pipelines

### OOXML: DOCX, XLSX, PPTX

These are OPC packages: ZIP containers with relationships, content types, and XML or binary parts. A reader must safely locate package entries, parse relationships without following external targets, and then interpret format-specific parts. This implementation supplies a bounded ZIP reader and a small XML parser, followed by separate Word, workbook, and presentation model builders. The parsers ignore external relationships for resource loading. Relationship targets are normalized within their package directory.

XLSX resolves workbook themes, cell and differential styles, number formats, shared and inline strings, defined names, tables, slicers, sparklines, comments, hyperlinks, data-validation prompts, and selected drawings. Rendering uses an absolutely positioned worksheet grid with dimension handling and frozen panes. Drawings share the DrawingML renderer with PPTX. Automated tests use synthetic ZIP/XML fixtures and exercise parser output and selected calculations; they do not establish browser visual equivalence with Excel.

DOCX additionally walks content controls and common wrapped-content elements, resolves basic `basedOn` style chains and document defaults, reads numbering definitions, section dimensions and margins, and follows internal references to headers, footers, and supported images. Rendering uses browser HTML and SVG with explicit font, paragraph, table, and page dimensions. Basic anchored DrawingML groups with custom paths and theme colors are rendered behind or alongside text. Explicit page breaks begin another page surface; automatic line-based pagination and full Word layout are not implemented.

Relevant references:

- ECMA-376, Office Open XML: <https://ecma-international.org/publications-and-standards/standards/ecma-376/>
- ISO/IEC 29500, Office Open XML (standard listing): <https://www.iso.org/standard/71691.html>
- W3C XML 1.0: <https://www.w3.org/TR/xml/>
- W3C Namespaces in XML: <https://www.w3.org/TR/xml-names/>
- PKWARE ZIP Application Note: <https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT>
- RFC 1951, DEFLATE: <https://www.rfc-editor.org/rfc/rfc1951>

### OpenDocument spreadsheets: ODS, OTS, FODS

An ODS or OTS file is a ZIP package whose first entry, `mimetype`, names the document type, with a manifest in `META-INF/manifest.xml` and XML parts for content, styles, metadata, and settings. A FODS file holds the same content in one XML document. The package is validated before any part is parsed: an encrypted document is refused, and signatures and macros are reported but never verified or run. Element and attribute names are matched by namespace URI, so any prefix a producer chooses works.

Cell values come only from the value attributes, never from the display text. The text the producing application saved is what the grid shows. Number styles are converted to the shared number-format engine and formatted in the document's own locale, never the viewer's. That formatted text is used only where no text was saved, and to check saved text (a difference is reported, and the saved text is still shown). Styles that cannot be reproduced exactly, such as embedded text, fill characters, non-Gregorian calendars, General, optional decimals, and fixed fraction denominators, are left to the saved text. Formulas are kept as text with their saved results; this library does not evaluate OpenFormula.

Hidden or clipped data is always reported, never silently dropped. The workbook notice above the grid lists accuracy and warning findings (cached formula results, hidden rows, columns and sheets, unformatted values, mismatched saved text, and unsupported charts, drawings, conditional formatting, and external links). Informational findings are kept in the parsed model only.

References:

- OASIS Open Document Format for Office Applications (OpenDocument) Version 1.3: <https://docs.oasis-open.org/office/OpenDocument/v1.3/OpenDocument-v1.3-part1-introduction.html>
- OASIS OpenDocument Version 1.3 Part 2 (packages) and Part 3 (schema): <https://docs.oasis-open.org/office/OpenDocument/v1.3/>

### ZIP and XML foundations

The ZIP reader scans the permitted end-record window, reads central-directory entries, supports stored and raw-DEFLATE data, verifies CRC-32 and declared sizes, and caps entry count and expansion size. It rejects encrypted, multi-disk, ZIP64, unknown-method, duplicate-path, and traversal-path entries. The custom DEFLATE implementation handles stored, fixed-Huffman, and dynamic-Huffman blocks. ZIP64 and data descriptors are only usable when central-directory metadata is present; ZIP64 itself is rejected.

The XML parser supports nested elements, qualified names, attributes, text, comments, CDATA, processing instructions, predefined and numeric entities, and self-closing tags. DTDs and entity declarations are rejected, and there is no external resource resolution. Namespace prefixes are retained on attributes and exposed as local element names. XML canonicalization, validation, and schema processing are not implemented.

OpenDocument parts use a separate namespace-aware tokenizer and streaming reader. They resolve prefixes to namespace URIs, apply the same DTD and entity rules, and hold only bounded subtrees (style blocks and one table cell at a time) in memory, so large spreadsheets are not built as a full XML tree.

### PDF

PDF signature, MIME, and extension detection exist, but there is no PDF parser or renderer. A custom renderer can be registered by an application. Detection alone must not be presented as built-in PDF support.

References:

- ISO 32000-2 (PDF 2.0): <https://www.iso.org/standard/75839.html>
- Adobe PDF 1.7 reference (public specification): <https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/PDF32000_2008.pdf>

### Legacy Office: DOC, XLS, PPT

Legacy files use Compound File Binary containers and distinct Word, BIFF, and PowerPoint record formats; they cannot be interpreted as OOXML. The implementation plan is a shared bounded CFB reader followed by separate parsers. There is no current parser or renderer for these formats.

References:

- Microsoft Open Specifications, [MS-CFB](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/)
- Microsoft Open Specifications, [MS-DOC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/)
- Microsoft Open Specifications, [MS-XLS](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/)
- Microsoft Open Specifications, [MS-PPT](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/)

### GIF and images

PNG, JPEG, WebP, and GIF previews use browser-native image decoding. The initial GIF path delegates only image decoding to the browser; this package does not independently model GIF frames, disposal, delays, or loop metadata. SVG input is parsed by the project's XML parser (files with a DOCTYPE are refused) and sanitized to an allowlist before serialization into an image context. Script elements, event attributes, style attributes, external references, and unrecognized elements/attributes are removed.

References:

- GIF89a specification: <https://www.w3.org/Graphics/GIF/spec-gif89a.txt>
- SVG 2: <https://www.w3.org/TR/SVG2/>

## Security and limits

- Default file limit is 128 MiB; ZIP entries default to 32 MiB and the package-entry count defaults to 4,096.
- The XML parser bounds source size, element count, and nesting depth; it never resolves external entities.
- The ZIP implementation rejects unsupported/encrypted/multi-disk/ZIP64 content rather than guessing.
- OpenDocument text documents have their own limits: `content.xml` (or a whole `.fodt` file) 64 MiB, `styles.xml` 16 MiB, all package parts together 256 MiB, 100,000 blocks (paragraphs, headings and tables), 100,000 table cells, 1,000,000 characters in one paragraph, 10,000 elements in one captured paragraph, and 32 MiB for one image stored inline in a flat document. The XML elements of the content stream are bounded by the block budget: 250,000 plus 24 per permitted block, so 2,650,000 at the default limit, and four tokens per element.
- OpenDocument spreadsheets have their own limits: `content.xml` (or a whole `.fods` file) 128 MiB, `styles.xml` 16 MiB, all package parts together 256 MiB, 500,000 cells, 1,024 sheets, 32,767 characters of text per cell, and 10,000 elements per captured cell or style block. The XML elements of the content stream are bounded by the cell limit rather than by the permitted part size: 250,000 plus eight per permitted cell, which is 4,250,000 at the default cell limit, counting every element read including those the parser discards. Namespace resolution copies nothing: entering an element costs only its own declarations, which also covers SVG. Exceeding a limit stops the preview with a `resource-limit` error; content is never truncated silently. Encrypted (password-protected) OpenDocument files are refused before any part is parsed.
- OpenDocument formulas are never evaluated, external data (linked tables, DDE) is never fetched, macros are not run, signatures are not verified, and only `http`, `https`, and `mailto` cell links are kept.
- XLSX cached formula values are shown as saved. A small built-in evaluator (no `eval`, no external links, no I/O) computes conditional-format rules and formula cells saved without a cached value; unsupported functions leave the cell marked as not calculated.
- Embedded scripts are not executed. External OOXML relationships are ignored.
- Embedded images are shown only as PNG, JPEG, GIF, BMP, WebP, or sanitized SVG, with the type taken from the content's signature; mislabelled media is not shown.
- DOCX, PPTX, and XLSX packages have a 256 MiB total decompression budget, each part is read once, and XLSX calculation, hyperlink ranges, and images have work limits; see [Security](./security.md).
- Browser-native image decoding still runs in the browser process; this package cannot promise isolation from browser codec vulnerabilities.
- Cancellation is supported for fetch-based sources, lifecycle replacement, and `load(file, { signal })`. A parser executing on the main thread cannot be preempted; in worker mode a running parse is stopped by terminating the worker, and a timeout ends parses that take too long.

## Test evidence and status

The current suite contains about 340 Node.js tests, primarily synthetic parser fixtures, characterization snapshots of the OOXML parser output, and low-level ZIP/XML/binary tests. A separate, manual browser regression script (`scripts/browser-regression.mjs`) compares the rendered DOM, stylesheet, and screenshots of the demo samples and synthetic OOXML fixtures against a baseline recorded from an earlier commit; it needs a local Edge or Chrome and is not part of `npm test` (CI runs it, see below). With `--worker` it renders every case in worker mode, so a worker-mode run can be compared with a main-thread baseline. A second manual script (`scripts/browser-security.mjs`) renders malicious fixtures and checks that embedded SVG runs no script and that font names cannot load remote resources; it also renders every fixture under the recommended Content Security Policy and under Trusted Types, in main-thread and worker mode, and checks that the worker timeout stops a slow parse while the page stays responsive. CI runs both browser scripts in headless Chrome on every push, a denial-of-service corpus with time and memory budgets (`npm run test:security`), and a zero-dependency mutation fuzzer over the parsers (`scripts/fuzz.mjs`; one minute per push, twenty minutes nightly). There are no included real-world documents. The tests validate only the parsers, algorithms, and samples they exercise; they do not demonstrate compatibility with arbitrary files or visual equivalence to Word, Excel, PowerPoint, or LibreOffice. The renderers should be treated as experimental subsets.
