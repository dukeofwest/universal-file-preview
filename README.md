# Universal File Preview

Universal File Preview is a browser-side TypeScript library for previewing common image formats and selected content from Office Open XML documents. It provides file loading, format detection, a typed viewer lifecycle, and a renderer registry. Its ZIP, DEFLATE, XML, document parsing, and rendering code is implemented in this project; it does not depend on third-party document parsing or rendering packages.

> **Pre-release:** The package is not currently available from the public npm registry. The package name and version are configured as `universal-file-preview@0.1.0`, but npm did not return a public package record during release preparation. Do not rely on `npm install universal-file-preview` until publication is announced.

## Features

- Preview GIF, PNG, JPEG, WebP, and sanitized SVG using browser image decoding.
- Render implemented subsets of DOCX, XLSX, and PPTX in the browser.
- Load from a `File`/`Blob`, `ArrayBuffer`, `Uint8Array`, or a URL fetched by the browser.
- Replace the current file, zoom or fit the preview, download loaded bytes, and clean up renderer resources.
- Add format handlers through a renderer registry.
- Apply default viewer styles inside the viewer root; optionally import the package CSS entry.
- Bound input and archive resource use, reject unsafe ZIP paths, and return typed preview errors.

The supported-format matrix lists what is detected, parsed, rendered, and covered by automated tests. The Office renderers are partial implementations and do not claim pixel-perfect compatibility with Microsoft Office.

## Quick start

After the package is published, install it with npm, pnpm, or Yarn as described in [Installation](./docs/installation.md). The following example uses a browser file input and the public `FilePreviewer` API:

```ts
import { FilePreviewer } from "universal-file-preview";

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

## Supported formats

| Format | Current implementation |
| --- | --- |
| GIF, PNG, JPEG/JPG, WebP | Browser-native image decoding; SVG is sanitized before image decoding. |
| DOCX | Partial Word text, styles, tables, numbering, headers/footers, page setup, and selected images/drawings. No automatic pagination. |
| XLSX | Styled worksheet grid with many cell, number, drawing, conditional-format, table, and slicer features. Charts and many formulas are not supported. |
| PPTX | Slide text, layout/theme/master content, shapes, images, tables, backgrounds, and navigation, with Office fidelity gaps. |
| PDF, DOC, XLS, PPT | Detection only. No built-in parser or renderer. |

For detailed coverage, test status, and limitations, see [Supported formats](./docs/format-support.md).

## API and configuration

The main entry point is `FilePreviewer`. Its options include `container`, `file`, `toolbar`, `zoom`, `download`, `maxFileSize`, and `registry`. The viewer exposes `open()`, `load(file)`, `setZoom(value)`, `fitToWidth()`, `fitToPage()`, and `destroy()`. See [API reference](./docs/api-reference.md) for public exports and exact types.

The package injects its scoped styles by default. Import `universal-file-preview/styles.css` when the prefixed stylesheet classes are also needed outside the viewer root.

## Architecture

The flow is file source → bounded byte loading → signature/metadata detection → format parser → typed document model → registered renderer → DOM/SVG output. `FilePreviewer` owns the lifecycle and common controls. The ZIP/DEFLATE reader and XML parser are project code; image codecs and DOM/fetch APIs are supplied by the browser. See [Architecture](./docs/architecture.md).

## Browser and TypeScript requirements

This is a browser package and uses browser DOM, `Blob`, `URL.createObjectURL`, `fetch`, and related APIs. The emitted JavaScript target is ES2022. The project does not yet publish a tested browser-version compatibility matrix. Development CI currently uses Node.js 22; a minimum supported development Node version has not been declared.

## Security and privacy

Local file sources are read in the browser and are not uploaded by this library. URL sources are fetched by the browser and remain subject to CORS and your application’s network policies. SVG is restricted to an allowlist; external OOXML relationships are not followed. Inputs are resource-bounded, but synchronous document parsing runs on the main thread. Review [Security](./docs/security.md) before previewing untrusted files.

## Limitations

- DOCX, XLSX, and PPTX are partial implementations; visual and feature fidelity varies by file.
- There is no built-in PDF or legacy binary Office renderer.
- Browser rendering and real-file visual comparisons are not covered by the current automated test suite.
- Remote resources require CORS permission. ZIP64 and encrypted OOXML packages are unsupported.
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

MIT. See [LICENSE](./LICENSE). Confirm that this license and the included copyright notice match the project owner’s intent before public distribution.

## Documentation

- [Installation](./docs/installation.md)
- [Getting started and examples](./docs/getting-started.md)
- [API reference](./docs/api-reference.md)
- [Supported formats](./docs/format-support.md)
- [Architecture](./docs/architecture.md)
- [Security](./docs/security.md)
- [Troubleshooting](./docs/troubleshooting.md)
- [Development and testing](./docs/development.md)
- [Release and publishing](./docs/release-and-publishing.md)

No GitHub repository URL or issue tracker is configured in this checkout, so no repository links are listed yet.
