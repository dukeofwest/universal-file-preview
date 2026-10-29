# API reference

The root entry point is ESM-only. Public exports are defined in `src/index.ts`. Unless noted otherwise, browser DOM APIs must be available before constructing a viewer.

## `FilePreviewer`

```ts
new FilePreviewer(options: PreviewOptions)
```

Creates a viewer and appends its DOM and scoped styles to `options.container`.

### `PreviewOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `container` | `HTMLElement` | required | Parent that receives the viewer. |
| `file` | `PreviewFile` | required | Initial file metadata and source. |
| `toolbar` | `boolean` | `true` | Show or hide the toolbar. |
| `zoom` | `boolean` | `true` | Show or hide zoom and fit controls. |
| `download` | `boolean` | `true` | Show or hide the download button. |
| `maxFileSize` | `number` | `128 * 1024 * 1024` | Maximum input size in bytes; must be a positive safe integer. |
| `registry` | `RendererRegistry` | default registry | Renderer registry to use. |

### `PreviewFile`

```ts
interface PreviewFile {
  readonly name: string;
  readonly source: string | Blob | ArrayBuffer | Uint8Array;
  readonly mimeType?: string;
}
```

String sources are fetched with browser `fetch`. Binary sources are read in memory. `name` is required and is used in the UI, downloads, and format-detection fallback.

### Methods

| Method | Return | Behavior |
|---|---|---|
| `open()` | `Promise<void>` | Loads the constructor's initial file. |
| `load(file: PreviewFile)` | `Promise<void>` | Aborts/replaces the previous load and renders the new file. Errors reject with `PreviewError` for viewer and parser failures. |
| `setZoom(value: number)` | `void` | Sets zoom; non-finite values throw `RangeError`. The viewer clamps the value to 0.1–4, and individual renderers may apply narrower limits. |
| `fitToWidth()` | `void` | Asks the current renderer to fit the preview to the available width. |
| `fitToPage()` | `void` | Currently aliases `fitToWidth()`. |
| `destroy()` | `void` | Idempotently aborts loading, calls renderer cleanup, revokes owned download resources, and removes viewer DOM/styles. |

Calls to `open`, `load`, or zoom methods after destruction throw `PreviewError` with code `renderer-error`. There is no public callback/event API for load status; the component updates its own status region and its promise reports errors.

## Errors

`PreviewError extends Error` and exposes a readonly `code`:

| Code | Meaning |
|---|---|
| `unsupported-format` | A format was detected but no renderer is registered. |
| `invalid-file` | Input or a supported package part is malformed. |
| `resource-limit` | A configured or built-in size/count bound was exceeded. |
| `network-error` | A remote request failed or returned a non-success status. |
| `parse-error` | An unexpected parser or renderer error was wrapped. |
| `unsupported-feature` | The parser encountered a deliberately unsupported construct. |
| `renderer-error` | Renderer setup/registration/lifecycle failed or the viewer is already destroyed. |
| `aborted` | A load was cancelled. A replaced stale load may finish without changing the new preview. |

## Renderer extension API

```ts
interface RendererContext {
  readonly file: PreviewFile;
  readonly format: FileFormat;
  readonly bytes: Uint8Array;
  readonly signal: AbortSignal;
}

interface RendererInstance {
  readonly element: HTMLElement;
  setZoom?(zoom: number): void;
  fitToWidth?(): void;
  destroy?(): void;
}

interface FileRenderer {
  readonly formats: readonly FileFormat[];
  render(context: RendererContext): RendererInstance | Promise<RendererInstance>;
}
```

`RendererRegistry.register(renderer)` registers every format in the renderer's `formats` list and returns an unregister function. Duplicate formats in one renderer or formats already registered cause `PreviewError("renderer-error")`. `createDefaultRendererRegistry()` returns the built-in image/DOCX/XLSX/PPTX registry.

## Format and lower-level utilities

- `detectFormat(name, mimeType?, bytes?)` returns `FileFormat`, including `"unknown"`. Detection is not proof that a built-in renderer exists.
- `BinaryReader` provides checked byte reads, seeks, endian-aware integer reads, and text reads.
- `ZipArchive.open(bytes, { maxEntrySize?, maxEntries? })` reads supported ZIP metadata; `entries`, `has(path)`, and `read(path)` expose package entries.
- `inflateRaw(bytes, expectedSize, maxOutputSize?)` exposes the project's raw-DEFLATE decoder; the default maximum output is 32 MiB.
- `parseXml`, `elements`, `firstElement`, and `textContent` expose the project's bounded XML subset.
- `parseDocx`, `parseXlsx`, and `parsePptx` expose low-level parser functions and typed models. These are parser utilities, not a promise that every input feature is supported.

The root entry point also exports these model types:

- Word and shared text: `Paragraph`, `TextRun`, `Table`, `TableCell`, `WordBlock`, `WordDocument`, `WordPageLayout`.
- Spreadsheet: `SpreadsheetAlignment`, `SpreadsheetAnchor`, `SpreadsheetBorderEdge`, `SpreadsheetBorders`, `SpreadsheetCell`, `SpreadsheetColumn`, `SpreadsheetConditional`, `SpreadsheetDrawing`, `SpreadsheetFill`, `SpreadsheetFont`, `SpreadsheetRange`, `SpreadsheetRow`, `SpreadsheetRun`, `SpreadsheetSlicer`, `SpreadsheetSlicerStyle`, `SpreadsheetSparkline`, `SpreadsheetStyle`, `SpreadsheetTable`, `Workbook`, `Worksheet`.
- Presentation: `Presentation`, `PptxCamera`, `PptxElement`, `PptxFill`, `PptxImageElement`, `PptxLine`, `PptxLineElement`, `PptxLineEnd`, `PptxReflection`, `PptxShapeElement`, `PptxTableElement`, `PptxTextElement`.

Model fields can evolve as format coverage changes and should not be assumed to be more stable than the viewer API.

## CSS subpath

The `universal-file-preview/styles.css` export provides the prefixed stylesheet. The viewer itself injects styles needed for its internal presentation, so applications do not need a global CSS import to render a preview.
