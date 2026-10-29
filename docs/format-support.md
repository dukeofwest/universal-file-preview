# Format support and implementation matrix

This document records the current implementation rather than promising complete compatibility. Parsers treat input as untrusted and return an error when the implemented subset cannot be read safely.

| Extension(s) | Detection | Parsing | Rendering | Implemented subset and material gaps | Automated tests / status |
|---|---|---|---|---|---|
| `.png`, `.jpg`, `.jpeg`, `.webp`, `.gif` | Signature, MIME, or extension | Browser handles image decoding; no project image codec | Yes, via browser `<img>` decoding | GIF animation is browser-provided. Codec support depends on the browser. | Detection has synthetic tests; browser decoding and viewer presentation are not automated in this suite. Partial. |
| `.svg` | XML root, MIME, or extension | Browser XML DOM plus project allowlist sanitizer | Yes, via browser `<img>` decoding | Only allowlisted elements and attributes survive. Scripts, event handlers, style attributes, external references, and unsupported SVG features are removed. | Detection is tested; sanitizer and browser presentation have no dedicated automated tests. Partial. |
| `.docx` | ZIP signature plus extension/MIME | Yes, bounded OOXML package and Word subset | Yes, DOM/SVG renderer | Paragraph/run styles, numbering, tables, page dimensions/margins, headers/footers, selected images, and basic anchored DrawingML. No automatic pagination; incomplete floating layout, fields, comments, footnotes, tracked changes, advanced table-style inheritance, and full Word typography. | Four synthetic parser fixtures; no automated browser rendering or visual-regression tests. Experimental subset. |
| `.xlsx` | ZIP signature plus extension/MIME | Yes, workbook and worksheet subset | Yes, styled, absolutely positioned worksheet grid | Cell styles, number formats, rich text, dimensions, merges, frozen panes, conditional formatting, tables, slicers, sparklines, notes, web links, selected drawings, TIFF conversion, and limited formula evaluation. Charts, many formulas, right-to-left sheets, outline controls, and some effects/styles are unsupported or approximated. | Eight synthetic parser/algorithm fixtures across `xlsx.test.mjs` and `ooxml.test.mjs`; no automated browser rendering or visual-regression tests. Experimental subset. |
| `.pptx` | ZIP signature plus extension/MIME | Yes, presentation, slide, theme, layout, master, and selected drawing subset | Yes, slide renderer with keyboard and previous/next navigation | Text, backgrounds, theme references, placeholders, many shapes/lines, images, tables, selected SmartArt drawing content, gradients, and some effects. No full PowerPoint fidelity; charts, animation, transitions, and embedded media are not implemented. | Two synthetic parser fixtures; no automated browser rendering or real-file visual regression. Experimental subset. |
| `.pdf` | Signature, MIME, or extension | No | No built-in renderer | Detection only. | Detection covered by a synthetic test. Unsupported for preview without a custom renderer. |
| `.doc`, `.xls`, `.ppt` | Extension/MIME; a CFB signature is only used with subtype metadata | No CFB or legacy format parser | No built-in renderer | Detection only; these are not equivalent to their OOXML counterparts. | Detection covered by a synthetic test. Unsupported for preview without a custom renderer. |

Detection does not imply the package can render a file. The default renderer registry contains handlers only for the image formats, DOCX, XLSX, and PPTX rows marked as implemented above. A detected PDF or legacy Office file produces an `unsupported-format` error unless the application registers a custom renderer.

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

### ZIP and XML foundations

The ZIP reader scans the permitted end-record window, reads central-directory entries, supports stored and raw-DEFLATE data, verifies CRC-32 and declared sizes, and caps entry count and expansion size. It rejects encrypted, multi-disk, ZIP64, unknown-method, duplicate-path, and traversal-path entries. The custom DEFLATE implementation handles stored, fixed-Huffman, and dynamic-Huffman blocks. ZIP64 and data descriptors are only usable when central-directory metadata is present; ZIP64 itself is rejected.

The XML parser supports nested elements, qualified names, attributes, text, comments, CDATA, processing instructions, predefined and numeric entities, and self-closing tags. DTDs and entity declarations are rejected, and there is no external resource resolution. Namespace prefixes are retained on attributes and exposed as local element names. XML canonicalization, validation, and schema processing are not implemented.

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

PNG, JPEG, WebP, and GIF previews use browser-native image decoding. The initial GIF path delegates only image decoding to the browser; this package does not independently model GIF frames, disposal, delays, or loop metadata. SVG input is parsed by the browser's XML DOM parser and sanitized to an allowlist before serialization into an image context. Script elements, event attributes, style attributes, external references, and unrecognized elements/attributes are removed.

References:

- GIF89a specification: <https://www.w3.org/Graphics/GIF/spec-gif89a.txt>
- SVG 2: <https://www.w3.org/TR/SVG2/>

## Security and limits

- Default file limit is 128 MiB; ZIP entries default to 32 MiB and the package-entry count defaults to 4,096.
- The XML parser bounds source size, element count, and nesting depth; it never resolves external entities.
- The ZIP implementation rejects unsupported/encrypted/multi-disk/ZIP64 content rather than guessing.
- XLSX cached formula values are shown as saved. A small built-in evaluator (no `eval`, no external links, no I/O) computes conditional-format rules and formula cells saved without a cached value; unsupported functions leave the cell marked as not calculated.
- Embedded scripts are not executed. External OOXML relationships are ignored.
- Browser-native image decoding still runs in the browser process; this package cannot promise isolation from browser codec vulnerabilities.
- Cancellation is supported for fetch-based sources and lifecycle replacement. Synchronous parsing cannot be preempted while a parser is executing on the main thread.

## Test evidence and status

The current suite contains 20 Node.js tests, primarily synthetic parser fixtures and low-level ZIP/XML/binary tests. There are no automated DOM/browser rendering tests, visual snapshots, or included real-world documents. The `tests/` suite therefore validates only the parser and algorithms it directly exercises; it does not demonstrate compatibility with arbitrary files or visual equivalence to Word, Excel, or PowerPoint. The renderers should be treated as experimental subsets.
