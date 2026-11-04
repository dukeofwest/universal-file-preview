# Angular integration example

A small standalone Angular 22 application using the published `@dukeofwest/universal-file-preview` package (`0.2.0`). The example uses strict TypeScript and a standalone component. It does not import the library source from this repository.

## Requirements

- Node.js 22.22.3 or newer on the 22.x line (or 24.15 or newer) and npm
- A browser with DOM and `Blob` support

## Install and run

From this directory:

```sh
npm install
npm start
```

Open the local URL printed by Angular CLI. Use **Choose file** or drop a supported image, SVG, DOCX, XLSX, PPTX, ODS, OTS, or FODS file. The browser reads the selected file locally; the demo does not upload it. The previewer's toolbar includes zoom and original-file download controls. Unsupported formats and parse failures show an error.

## How it works

`AppComponent` is standalone and imports the public `FilePreviewer` class from the npm package. It creates the viewer when a file is selected, passes the browser `File` as the `PreviewFile.source`, and calls `destroy()` when the selected file is cleared or the component is removed. `ngOnDestroy` releases viewer resources. The file input and drag-and-drop handlers both call the same typed `previewFile` method.

The essential integration is:

```ts
import { FilePreviewer } from "@dukeofwest/universal-file-preview";

const preview = new FilePreviewer({
  container: previewHost,
  file: { name: file.name, source: file, mimeType: file.type },
});
await preview.open();
// In ngOnDestroy, or before replacing/removing this viewer:
preview.destroy();
```

The component catches the rejected `open()` promise and shows its message. The package also displays its own loading, unsupported-format, and parse-error status in the preview area.

## Build

```sh
npm run build
npm run typecheck
```

The production build is configured with a relative base href and writes to the shared Pages staging directory when invoked from the root workflow.
