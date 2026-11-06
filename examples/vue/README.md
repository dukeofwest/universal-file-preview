# Vue 3 integration example

A small Vue 3 application using the published `@dukeofwest/universal-file-preview` package (`0.2.2`). It uses the Composition API, strict TypeScript checks, and does not import the library source from this repository.

## Requirements

- Node.js 22.22.3 or newer on the 22.x line (or 24.15 or newer) and npm
- A browser with DOM and `Blob` support

## Install and run

From this directory:

```sh
npm install
npm run dev
```

Open the local URL printed by Vite. Choose or drop a supported image, SVG, DOCX, XLSX, PPTX, ODS, OTS, or FODS file. The browser reads it locally; this demo does not upload it. The viewer toolbar includes zoom and original-file download controls. Unsupported formats and parse failures show an error.

## How it works

`src/App.vue` imports `FilePreviewer` from the package and holds it in a `shallowRef`. This avoids deeply proxying a class instance. When a file is selected, the app creates a previewer with the host element and the public `PreviewFile` shape `{ name, source, mimeType }`, then awaits `open()`. It destroys the old viewer on replacement, clear, and `onBeforeUnmount`. The file input and drop target share the same function.

The core lifecycle is:

```ts
import { onBeforeUnmount, shallowRef } from "vue";
import { FilePreviewer } from "@dukeofwest/universal-file-preview";

const previewer = shallowRef<FilePreviewer | null>(null);
onBeforeUnmount(() => previewer.value?.destroy());

async function showFile(file: File, container: HTMLElement): Promise<void> {
  previewer.value?.destroy();
  const instance = new FilePreviewer({
    container,
    file: { name: file.name, source: file, mimeType: file.type },
  });
  previewer.value = instance;
  await instance.open();
}
```

The example catches `open()` errors and shows them above the viewer. The package renders loading and format errors inside its preview area too.

## Build

```sh
npm run build
npm run preview
```

The Vite production build uses relative asset URLs and writes to the shared Pages staging directory when invoked from the root workflow.
