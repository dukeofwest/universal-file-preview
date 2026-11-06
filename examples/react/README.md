# React integration example

A small React 19 application using the published `@dukeofwest/universal-file-preview` package (`0.2.2`). It uses function components, hooks, strict TypeScript, and does not import the library source from this repository.

## Requirements

- Node.js 22.22.3 or newer on the 22.x line (or 24.15 or newer) and npm
- A browser with DOM and `Blob` support

## Install and run

From this directory:

```sh
npm install
npm run dev
```

Open the local URL printed by Vite. Choose or drop a supported image, SVG, DOCX, XLSX, PPTX, ODS, OTS, or FODS file. The browser reads it locally; the demo does not upload it. The viewer toolbar includes zoom and original-file download controls. Unsupported formats and parse failures show an error.

## How it works

`src/App.tsx` creates a `FilePreviewer` when a `File` is selected, uses the DOM host held by a React ref as `container`, and passes the file in the documented `{ name, source, mimeType }` shape. It destroys the previous preview before opening another and destroys the active preview in the effect cleanup. Both the file input and drop target use the same preview function.

The core lifecycle is:

```tsx
import { useEffect, useRef } from "react";
import { FilePreviewer } from "@dukeofwest/universal-file-preview";

const host = useRef<HTMLDivElement>(null);
const previewer = useRef<FilePreviewer | null>(null);
useEffect(() => () => previewer.current?.destroy(), []);

async function showFile(file: File): Promise<void> {
  const container = host.current;
  if (!container) return;
  previewer.current?.destroy();
  const instance = new FilePreviewer({
    container,
    file: { name: file.name, source: file, mimeType: file.type },
  });
  previewer.current = instance;
  await instance.open();
}
```

The example catches `open()` errors and displays them above the preview area. The package also renders loading and format errors inside its viewer.

## Build

```sh
npm run build
npm run preview
```

The Vite production build uses relative asset URLs and writes to the shared Pages staging directory when invoked from the root workflow.
