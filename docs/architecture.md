# Architecture

## Runtime flow

`FilePreviewer` owns the viewer DOM, toolbar, loading status, current file, zoom state, download action, cancellation, and cleanup. It obtains bytes from a `Blob`, `ArrayBuffer`, `Uint8Array`, or browser `fetch` request. It detects a `FileFormat`, finds the matching handler in `RendererRegistry`, and passes the file metadata, format, bytes, and an `AbortSignal` to that renderer. The renderer returns a DOM element and may implement zoom, fit-to-width, and cleanup methods.

The built-in registry contains image, DOCX, XLSX, PPTX, and OpenDocument spreadsheet (ODS, OTS, FODS) handlers. PDF and legacy Office formats can be detected but have no built-in renderer. Applications can register custom handlers.

## Parsing and rendering

### Images

PNG, JPEG, WebP, and GIF bytes are wrapped in a `Blob` and displayed with an `<img>` element; decoding is provided by the browser. SVG is decoded as text, parsed with `DOMParser`, reduced to an element/attribute allowlist, serialized, and then displayed as an image. This is not a general SVG implementation.

### Office Open XML

DOCX, XLSX, and PPTX are ZIP-based OPC packages. The project reads the ZIP central directory, validates package paths and entry metadata, inflates supported raw-DEFLATE entries, and parses XML into a small internal node structure. Format parsers follow internal relationships and build typed models. Renderers turn those models into HTML, CSS, and SVG.

The ZIP and DEFLATE readers, XML parser, OPC relationship handling, document parsers, and renderers are project code. They are partial implementations of the relevant file formats and are not complete Office-compatible engines. See [format support](./format-support.md) for precise limitations and test evidence.

### OpenDocument spreadsheets

ODS and OTS packages are read through a package reader that validates the ODF package contract (the `mimetype` entry, manifest, encryption, and per-part and total size budgets) before any part is parsed. `styles.xml`, `meta.xml`, and `settings.xml` are small and parsed into namespace-aware trees. `content.xml` is streamed: style blocks and one table cell at a time are captured as small trees, and rows are built as they arrive. A flat `.fods` file goes through the same reader in a single pass. The parser produces the same workbook model as XLSX, extended with typed values, display-text provenance, and parse diagnostics, and the XLSX worksheet grid renders both. Diagnostics of accuracy or warning severity appear in a notice above the grid.

### XLSX drawing reuse

Selected spreadsheet drawings are converted to the shared DrawingML element model used by the PPTX renderer. Spreadsheet anchors determine where those elements are positioned over the worksheet grid. TIFF drawings can be converted by the project’s limited TIFF-to-PNG implementation when they fit its supported subset.

## Resource lifecycle

Every `load()` cancels the previous fetch, destroys the current renderer instance, clears the content, and starts a new load generation. Stale asynchronous renderer results are discarded and cleaned up. `destroy()` aborts the active request, destroys renderer resources, revokes the viewer's download URL, and removes injected DOM and styles.

Image and embedded-image renderers create object URLs and revoke them on replacement or destruction. Applications must call `destroy()` when removing the viewer.

## Security boundaries

File bytes are treated as untrusted. Source size, ZIP entry size/count, package paths, XML size/element/depth, and selected calculation ranges are bounded in code. ZIP64, encryption, multi-disk ZIP, unsupported compression, encrypted OpenDocument packages, XML DTDs, and external entities are rejected. External OOXML relationships are not followed for embedded content. OOXML packages are read through `package/ooxml-package.ts`, which enforces a total decompression budget and reads and parses each part at most once. Every image blob URL is created in `renderers/image-url.ts`, which checks the type against the content's signature and sanitizes all SVG, including SVG embedded in Office files. See [Security](./security.md).

Parsing is synchronous once bytes have loaded, so large files can block the main thread. The code does not execute document scripts or formula expressions with `eval`; the XLSX evaluator supports only a limited built-in function/operator set, and OpenDocument formulas are not evaluated at all.

## Dependencies and platform APIs

There are no runtime npm dependencies. TypeScript is used to compile source and emit declarations. Tests use Node's built-in test runner. At runtime, the library relies on browser APIs such as DOM, `Blob`, `URL.createObjectURL`, `fetch`, `AbortController`, `DOMParser`, and native image decoding. The library has no worker requirement or worker asset.

## Extension points

`RendererRegistry.register()` adds a `FileRenderer` for one or more detected formats and returns a function that unregisters that renderer. A renderer receives the original input metadata, detected format, loaded bytes, and signal. It returns an element and optional lifecycle methods. Custom renderers are application-owned; this package does not provide missing format parsers when a custom handler is registered.
