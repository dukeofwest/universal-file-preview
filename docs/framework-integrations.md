# Framework integrations

Universal File Preview is a browser-side TypeScript library with a framework-agnostic DOM lifecycle API. It does not provide official Angular, React, or Vue adapter packages. The repository examples demonstrate how to own the viewer lifecycle from a standalone Angular component, React hooks, and Vue 3 Composition API.

All three examples install the published `@dukeofwest/universal-file-preview@0.2.0` package. They import the public `FilePreviewer` class, pass a framework-managed host element as `container`, and provide a browser `File` through the public `PreviewFile` shape. They destroy the current instance on cleanup and before replacing it.

```ts
import { FilePreviewer } from "@dukeofwest/universal-file-preview";

const preview = new FilePreviewer({
  container: previewElement,
  file: { name: file.name, source: file, mimeType: file.type },
});
await preview.open();
// Call when the host is removed, or before replacing the current file.
preview.destroy();
```

`FilePreviewer` also exposes `load(file)`, `setZoom(value)`, `fitToWidth()`, `fitToPage()`, and `destroy()`. The toolbar is enabled by default and includes zoom and original-file download buttons. `open()` and `load()` reject with typed `PreviewError` values when loading fails or a built-in renderer does not support the detected format. The viewer also displays loading and error status in its own DOM.

## Examples

- [Angular standalone example](../examples/angular/README.md)
- [React hooks example](../examples/react/README.md)
- [Vue Composition API example](../examples/vue/README.md)
- [Live demo deployment](./live-demo.md)

## Supported content

The built-in renderers support images (GIF, PNG, JPEG, WebP, sanitized SVG) and implemented subsets of DOCX, XLSX, and PPTX. PDF and legacy DOC/XLS/PPT are detected but have no built-in renderer. The Office renderers are partial and do not claim pixel-perfect compatibility. See [the format support matrix](./format-support.md) for verified features, tests, and limitations.

The examples process selected files in the browser. The applications do not send file contents to an external server. URL sources passed to the library are fetched by the browser and remain subject to CORS.
