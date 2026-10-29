# Getting started

These examples use the public API emitted from `src/index.ts`. The npm import becomes usable after a release is published; until then, run the examples against the local build.

## Preview a local file

Create the viewer after a file is selected. This keeps the required initial file explicit:

```ts
import { FilePreviewer, PreviewError } from "@dukeofwest/universal-file-preview";

const container = document.querySelector<HTMLElement>("#viewer");
const input = document.querySelector<HTMLInputElement>("#file");
if (!container || !input) throw new Error("Viewer elements are missing.");

let preview: FilePreviewer | undefined;
input.addEventListener("change", async () => {
  const file = input.files?.[0];
  if (!file) return;

  preview?.destroy();
  preview = new FilePreviewer({
    container,
    file: { name: file.name, source: file, mimeType: file.type },
  });
  try {
    await preview.open();
  } catch (error) {
    if (error instanceof PreviewError) console.error(error.code, error.message);
    else console.error(error);
  }
});

// Call from the component's unmount/dispose hook:
function disposePreview(): void {
  preview?.destroy();
  preview = undefined;
}
```

`FilePreviewer` requires an initial `PreviewFile`. `open()` loads that file; `load(file)` replaces it later. Loading and rendering errors are shown in the viewer status and are also rejected from the returned promise.

## Preview a Blob or byte buffer

```ts
const preview = new FilePreviewer({
  container,
  file: { name: "report.docx", source: fileBlob },
});
await preview.open();

await preview.load({ name: "workbook.xlsx", source: new Uint8Array(bytes) });
```

`source` also accepts an `ArrayBuffer`. Keep the correct filename extension when supplying raw bytes because the name is used as a fallback for format detection.

## Preview a remote file

```ts
const preview = new FilePreviewer({
  container,
  file: { name: "presentation.pptx", source: "https://files.example.test/deck.pptx" },
});
await preview.open();
```

The browser fetches the URL. The remote server must allow the application's origin with CORS. This package does not proxy the request.

## Replace a file and control the viewer

```ts
await preview.load({ name: nextFile.name, source: nextFile, mimeType: nextFile.type });
preview.setZoom(1.25);
preview.fitToWidth();
preview.fitToPage(); // Currently aliases fitToWidth().
```

Call `destroy()` when removing the viewer. It aborts an active fetch, calls the renderer cleanup hook, revokes object URLs owned by the viewer, and removes its DOM.

## Configure the viewer

```ts
const preview = new FilePreviewer({
  container,
  file: { name: file.name, source: file },
  toolbar: true,
  zoom: true,
  download: false,
  maxFileSize: 64 * 1024 * 1024,
});
await preview.open();
```

The default size limit is 128 MiB. `maxFileSize` is a positive safe integer in bytes. Setting `toolbar`, `zoom`, or `download` to `false` hides the corresponding controls.

## Register a custom renderer

```ts
import {
  createDefaultRendererRegistry,
  FilePreviewer,
  type FileRenderer,
} from "@dukeofwest/universal-file-preview";

const registry = createDefaultRendererRegistry();
const custom: FileRenderer = {
  formats: ["pdf"],
  async render({ file, bytes, signal }) {
    if (signal.aborted) throw new DOMException("Aborted", "AbortError");
    const element = document.createElement("div");
    element.textContent = `Application renderer for ${file.name} (${bytes.length} bytes)`;
    return { element, destroy: () => element.remove() };
  },
};
const unregister = registry.register(custom);

const preview = new FilePreviewer({
  container,
  file: { name: "document.pdf", source: pdfBlob },
  registry,
});
await preview.open();

// Dispose the viewer before removing the handler if it may still be active.
preview.destroy();
unregister();
```

The example registers an application-owned extension point; the built-in package does not parse or render PDF. Registering a second renderer for an already registered format throws `PreviewError` with code `renderer-error`.
