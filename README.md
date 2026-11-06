# Universal File Preview

[![npm version](https://img.shields.io/npm/v/%40dukeofwest%2Funiversal-file-preview)](https://www.npmjs.com/package/@dukeofwest/universal-file-preview)
[![MIT license](https://img.shields.io/github/license/dukeofwest/universal-file-preview)](./LICENSE)

**[Open the live Angular, React and Vue demo](https://dukeofwest.github.io/universal-file-preview/)** — try a framework integration with your own file or download a sample document.

Browser-based TypeScript previews for common image formats and selected content from DOCX, XLSX, PPTX, and OpenDocument spreadsheet (ODS) files. The library provides file loading, format detection, a typed viewer lifecycle, and a renderer registry. Its ZIP, DEFLATE, XML, document parsing, and rendering code is implemented in this project; it does not depend on third-party document parsing or rendering packages.

## Features

- Preview GIF, PNG, JPEG, WebP, and sanitized SVG using browser image decoding.
- Render implemented subsets of DOCX, XLSX, PPTX, and ODS/OTS/FODS in the browser.
- Load from a `File`/`Blob`, `ArrayBuffer`, `Uint8Array`, or a URL fetched by the browser.
- Replace the current file, zoom or fit the preview, download loaded bytes, and clean up renderer resources.
- Add format handlers through a renderer registry.
- Apply default viewer styles inside the viewer root; optionally import the package CSS entry.
- Limit file size, ZIP entry size and count, and XML size, reject unsafe ZIP paths, and return typed preview errors.

The supported-format matrix lists what is detected, parsed, rendered, and covered by automated tests. The Office renderers are partial implementations and do not claim pixel-perfect compatibility with Microsoft Office.

## Quick start

Install the package with `npm install @dukeofwest/universal-file-preview` as described in [Installation](./docs/installation.md). The following example uses a browser file input and the public `FilePreviewer` API:

```ts
import { FilePreviewer } from "@dukeofwest/universal-file-preview";

const container = document.querySelector<HTMLElement>("#viewer");
const input = document.querySelector<HTMLInputElement>("#file");
if (!container || !input) throw new Error("Viewer elements are missing.");

const preview = new FilePreviewer({
  container,
  file: { name: "welcome.svg", source: new Blob(["<svg></svg>"], { type: "image/svg+xml" }) },
});
await preview.open();

input.addEventListener("change", async () => {
  const file = input.files?.[0];
  if (!file) return;
  try {
    await preview.load({ name: file.name, source: file, mimeType: file.type });
  } catch (error) {
    console.error("Could not preview file", error);
  }
});

```

Call `preview.destroy()` from the owning UI's cleanup hook when the viewer is removed.

See [Getting started](./docs/getting-started.md) for remote files, byte buffers, error handling, switching files, and custom renderers.

## Angular, React and Vue

The framework examples are independent applications that consume the published npm package. They do not import the library source or add document-rendering dependencies. Each example uses a framework-managed element as the viewer container, passes a browser `File` to `FilePreviewer`, handles rejected loads, and calls `destroy()` when the viewer is cleared or the component is removed. Import the optional package stylesheet with `@dukeofwest/universal-file-preview/styles.css`.

Install the package in the framework application:

```sh
npm install @dukeofwest/universal-file-preview
```

The public integration shape is the same in each framework:

```ts
import { FilePreviewer } from "@dukeofwest/universal-file-preview";

const preview = new FilePreviewer({
  container: previewElement,
  file: { name: file.name, source: file, mimeType: file.type },
});
await preview.open();
// On clear, replacement, or component teardown:
preview.destroy();
```

In Angular, call this from a standalone component and destroy the instance in `ngOnDestroy`. In React, keep the host and instance in refs and destroy it from effect cleanup. In Vue 3, keep the instance in a `shallowRef` and destroy it in `onBeforeUnmount`. Catch errors from `open()` or `load()`; the viewer also displays its own loading and error status.

### Angular quick start

```ts
import { Component, ElementRef, OnDestroy, ViewChild } from "@angular/core";
import { FilePreviewer } from "@dukeofwest/universal-file-preview";

@Component({ standalone: true, template: `<div #host></div>` })
export class PreviewComponent implements OnDestroy {
  @ViewChild("host", { static: true }) host!: ElementRef<HTMLDivElement>;
  private preview: FilePreviewer | null = null;

  async openFile(file: File): Promise<void> {
    this.preview?.destroy();
    this.preview = new FilePreviewer({ container: this.host.nativeElement, file: { name: file.name, source: file, mimeType: file.type } });
    try { await this.preview.open(); } catch (error: unknown) { console.error("Preview failed", error); }
  }

  ngOnDestroy(): void { this.preview?.destroy(); }
}
```

### React quick start

```tsx
import { useEffect, useRef } from "react";
import { FilePreviewer } from "@dukeofwest/universal-file-preview";

function Preview() {
  const host = useRef<HTMLDivElement>(null);
  const preview = useRef<FilePreviewer | null>(null);
  useEffect(() => () => preview.current?.destroy(), []);

  async function openFile(file: File): Promise<void> {
    const container = host.current;
    if (!container) return;
    preview.current?.destroy();
    const instance = new FilePreviewer({ container, file: { name: file.name, source: file, mimeType: file.type } });
    preview.current = instance;
    try { await instance.open(); } catch (error: unknown) { console.error("Preview failed", error); }
  }

  return <div ref={host} />;
}
```

### Vue 3 quick start

```ts
import { onBeforeUnmount, ref, shallowRef } from "vue";
import { FilePreviewer } from "@dukeofwest/universal-file-preview";

const host = ref<HTMLDivElement | null>(null);
const preview = shallowRef<FilePreviewer | null>(null);
onBeforeUnmount(() => preview.value?.destroy());

async function openFile(file: File): Promise<void> {
  const container = host.value;
  if (!container) return;
  preview.value?.destroy();
  const instance = new FilePreviewer({ container, file: { name: file.name, source: file, mimeType: file.type } });
  preview.value = instance;
  try { await instance.open(); } catch (error: unknown) { console.error("Preview failed", error); }
}
```

- [Angular example and run instructions](./examples/angular/README.md)
- [React example and run instructions](./examples/react/README.md)
- [Vue example and run instructions](./examples/vue/README.md)
- [Framework integration guide](./docs/framework-integrations.md)

These are integration examples, not official framework adapter packages. See [format support](./docs/format-support.md) for the actual renderer coverage and limitations.

## Try the demo

The framework demo selector hosts the actual Angular, React, and Vue applications at [dukeofwest.github.io/universal-file-preview](https://dukeofwest.github.io/universal-file-preview/). The live site includes downloadable sample files so you can try each example without preparing your own documents.

The existing local library demo can still be run from a repository checkout:

```sh
npm ci
npm run build
npm run demo
```

Then open <http://127.0.0.1:4173/> and choose an image, DOCX, XLSX, PPTX, ODS, or FODS file.

See [live demo and Pages setup](./docs/live-demo.md) for deployment status and the local framework site preview instructions.

## Supported formats

| Format | Current implementation |
| --- | --- |
| GIF, PNG, JPEG/JPG, WebP | Browser-native image decoding; SVG is sanitized before image decoding. |
| DOCX | Partial Word text, styles, tables, numbering, headers/footers, page setup, and selected images/drawings. No automatic pagination. |
| XLSX | Styled worksheet grid with many cell, number, drawing, conditional-format, table, and slicer features. Charts and many formulas are not supported. |
| ODS, OTS, FODS | OpenDocument spreadsheets in the same worksheet grid: typed values with the producer's saved display text, styles, number formats, merges, hidden data, comments, links, and saved view settings. Formulas are not recalculated; charts, images, and conditional formatting are not shown, and a notice lists what is missing. |
| PPTX | Slide text, layout/theme/master content, shapes, images, tables, backgrounds, and navigation, with Office fidelity gaps. |
| PDF, DOC, XLS, PPT | Detection only. No built-in parser or renderer. |

For detailed coverage, test status, and limitations, see [Supported formats](./docs/format-support.md).

## API and configuration

The main entry point is `FilePreviewer`. Its options include `container`, `file`, `toolbar`, `zoom`, `download`, `maxFileSize`, and `registry`. The viewer exposes `open()`, `load(file)`, `setZoom(value)`, `fitToWidth()`, `fitToPage()`, and `destroy()`. See [API reference](./docs/api-reference.md) for public exports and exact types.

The package injects its scoped styles by default. Import `@dukeofwest/universal-file-preview/styles.css` when the prefixed stylesheet classes are also needed outside the viewer root.

## Architecture

The flow is file source → bounded byte loading → signature/metadata detection → format parser → typed document model → registered renderer → DOM/SVG output. `FilePreviewer` owns the lifecycle and common controls. The ZIP/DEFLATE reader and XML parser are project code; image codecs and DOM/fetch APIs are supplied by the browser. See [Architecture](./docs/architecture.md).

## Browser requirements and framework compatibility

This is a browser package and uses browser DOM, `Blob`, `URL.createObjectURL`, `fetch`, and related APIs. The emitted JavaScript target is ES2022. The project does not yet publish a tested browser-version compatibility matrix. Development CI currently uses Node.js 22; a minimum supported development Node version has not been declared.

The public API is framework-agnostic. The repository includes tested Angular, React, and Vue integration examples, but does not publish framework-specific adapter packages. Use the browser lifecycle API from your framework's mount and cleanup hooks. The examples build against the current published package version.

## Security and privacy

Local file sources are read in the browser and are not uploaded by this library. URL sources are fetched by the browser and remain subject to CORS and your application’s network policies. SVG, including SVG embedded in Office files, is restricted to an allowlist; external OOXML relationships are not followed. Inputs, decompression, calculation, images, and the parsed document have size and work limits. By default, document parsing runs on the main thread, so a large or crafted file can make the page unresponsive for several seconds; the optional worker mode parses documents in a Web Worker with a timeout instead. Read [Security](./docs/security.md), including its deployment guidance, before previewing untrusted files.

## Limitations

- DOCX, XLSX, PPTX, and ODS are partial implementations; visual and feature fidelity varies by file.
- ODT and ODP (OpenDocument text and presentations) are not supported yet.
- There is no built-in PDF or legacy binary Office renderer.
- Browser rendering and real-file visual comparisons are not covered by the current automated test suite.
- Remote resources require CORS permission. ZIP64 and encrypted OOXML and OpenDocument packages are unsupported.
- Font availability and image decoding depend on the browser and host system.

See [Troubleshooting](./docs/troubleshooting.md) and [Supported formats](./docs/format-support.md).

## Development and validation

```sh
npm ci
npm run build
npm test
npm run typecheck
npm pack --dry-run
```

Tests use Node’s built-in test runner and synthetic parser fixtures. TypeScript is a development-only dependency; there are no runtime dependencies. See [Development](./docs/development.md) and [Testing](./docs/testing.md).

## Contributing

Read [CONTRIBUTING.md](./CONTRIBUTING.md) before making changes. The implementation must continue using the project’s own document parsing and rendering engine; do not add third-party document parsers or renderers.

## License

MIT. See [LICENSE](./LICENSE).

## Documentation

- [Installation](./docs/installation.md)
- [Getting started and examples](./docs/getting-started.md)
- [Adoption and technical content plan](./docs/adoption.md)
- [Demo deployment](./docs/demo-deployment.md)
- [Framework integrations](./docs/framework-integrations.md)
- [Live framework demo](./docs/live-demo.md)
- [API reference](./docs/api-reference.md)
- [Supported formats](./docs/format-support.md)
- [Architecture](./docs/architecture.md)
- [Security](./docs/security.md)
- [Troubleshooting](./docs/troubleshooting.md)
- [Development and testing](./docs/development.md)
- [Release and publishing](./docs/release-and-publishing.md)

## Repository and support

- [GitHub repository](https://github.com/dukeofwest/universal-file-preview)
- [Issue tracker](https://github.com/dukeofwest/universal-file-preview/issues)
- [npm package](https://www.npmjs.com/package/@dukeofwest/universal-file-preview)
