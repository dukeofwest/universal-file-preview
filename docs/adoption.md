# Discoverability and adoption plan

This plan focuses on useful technical answers that help developers evaluate the package. It does not imply a feature is built in unless it appears in the [format support matrix](./format-support.md). Articles and external posts are proposals; none are published by this repository.

## Search queries to target

- How to preview local files in JavaScript or TypeScript.
- How to preview DOCX, XLSX, or PPTX files in a browser.
- Browser-based Office Open XML preview without uploading a file.
- How to load a remote document in a browser and handle CORS errors.
- How to add a custom file renderer to a TypeScript viewer.
- How to preview PDF in a custom browser application using a renderer extension (the package does not include a PDF renderer).

## Proposed technical articles

Each article should explain the supported subset and limitations, link to the relevant API and support matrix, and be tested against the published API before publication.

### 1. Preview a local file in a TypeScript web app

**Problem/query:** “How do I preview an uploaded image or Office document in the browser?”

Show a file input, the smallest `FilePreviewer` setup, typed error handling, and cleanup when the view is removed. Use this framework-neutral example:

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
  preview = new FilePreviewer({ container, file: { name: file.name, source: file, mimeType: file.type } });
  try {
    await preview.open();
  } catch (error) {
    if (error instanceof PreviewError) console.error(error.code, error.message);
    else console.error(error);
  }
});

function dispose(): void {
  preview?.destroy();
  preview = undefined;
}
```

### 2. Load a remote Office file and handle browser CORS

**Problem/query:** “Why can my server download a file when the browser cannot preview its URL?”

Explain that URL sources are fetched by the browser, so the file server must allow the application's origin. Demonstrate `PreviewError` without implying that the library proxies a request:

```ts
import { FilePreviewer, PreviewError } from "@dukeofwest/universal-file-preview";

const container = document.querySelector<HTMLElement>("#viewer");
if (!container) throw new Error("Viewer element is missing.");

const preview = new FilePreviewer({
  container,
  file: { name: "quarterly-report.xlsx", source: "https://files.example.test/quarterly-report.xlsx" },
});
try {
  await preview.open();
} catch (error) {
  if (error instanceof PreviewError && error.code === "network-error") {
    console.error("Check the response and CORS policy for the file server.");
  } else {
    console.error(error);
  }
}
```

### 3. Add a custom format renderer without replacing built-ins

**Problem/query:** “How can I connect my existing PDF renderer to a browser file viewer?”

Be explicit that this package detects PDF but does not parse or render it. The host application supplies and owns the PDF implementation:

```ts
import {
  createDefaultRendererRegistry,
  FilePreviewer,
  type FileRenderer,
} from "@dukeofwest/universal-file-preview";

const container = document.querySelector<HTMLElement>("#viewer");
if (!container) throw new Error("Viewer element is missing.");

const registry = createDefaultRendererRegistry();
const pdfRenderer: FileRenderer = {
  formats: ["pdf"],
  render({ file, bytes }) {
    const element = document.createElement("p");
    element.textContent = `Connect the host application's PDF renderer for ${file.name} (${bytes.length} bytes).`;
    return { element };
  },
};
const unregisterPdf = registry.register(pdfRenderer);
const preview = new FilePreviewer({
  container,
  file: { name: "document.pdf", source: new Uint8Array() },
  registry,
});

// After the host renderer has been connected and preview use has ended:
preview.destroy();
unregisterPdf();
```

The snippet demonstrates the extension contract only; its placeholder element is not a PDF viewer. An article that integrates a specific PDF library must keep that dependency in the sample application, not in this package, and clearly describe its licensing and security considerations.

### 4. Manage viewer replacement and cleanup in a single-page app

**Problem/query:** “How do I prevent old previews and object URLs from remaining after a user selects another document?”

Show one framework-neutral mount function that replaces the previous viewer and returns a cleanup function. Explain that framework adapters are not included or separately tested:

```ts
import { FilePreviewer, type PreviewFile } from "@dukeofwest/universal-file-preview";

function mountPreview(container: HTMLElement, file: PreviewFile): () => void {
  const preview = new FilePreviewer({ container, file });
  void preview.open().catch((error: unknown) => console.error("Preview failed", error));
  return () => preview.destroy();
}
```

## Distribution channels

- Publish the articles on Dev.to or Hashnode with runnable examples, links to the support matrix, and no unsupported fidelity claims.
- Share only the article that answers a specific question in relevant JavaScript/TypeScript communities, following each community's self-promotion rules.
- Use GitHub Discussions for feedback and support questions if Discussions are enabled; keep actionable defects in the issue tracker.
- Add relevant repository topics after reviewing the live repository settings: `file-preview`, `document-viewer`, `typescript`, `javascript`, `docx`, `xlsx`, `pptx`, and `open-source`.
- Do not repeatedly post package links or create accounts automatically. Ask for approval before publishing external content or changing repository settings.

After maintainer approval, the topics can be applied with GitHub CLI:

```sh
gh repo edit dukeofwest/universal-file-preview --add-topic file-preview --add-topic document-viewer --add-topic typescript --add-topic javascript --add-topic docx --add-topic xlsx --add-topic pptx --add-topic open-source
```

## Prioritized 30-day plan

| Timing | Work | Evidence of progress |
|---|---|---|
| Days 1–3 | Keep npm metadata, README, install instructions, and format support claims consistent with the currently published package. Add relevant GitHub topics after owner review. | Search result snippet and package page describe the same real features and limitations. |
| Days 4–7 | Publish article 1 with a small, runnable browser example. Link to it from the README only after the external page is live and reviewed. | A reader can install the package, select a local file, preview it, and clean up the viewer. |
| Days 8–14 | Publish article 2 about remote URLs and CORS; collect reproducible questions through GitHub. | Troubleshooting questions link to exact CORS and error behavior. |
| Days 15–21 | Publish article 3 about renderer extension points, especially the distinction between PDF detection and PDF rendering. | Extension example compiles against the release API and makes the host responsibility clear. |
| Days 22–30 | Publish article 4 about cleanup and lifecycle; review search queries and feedback, then refine documentation based on actual confusion. | A short list of user questions and documentation changes; no unsupported adoption or download claims. |

Do not report views, downloads, conversion, or compatibility metrics unless they have been measured from a real source and the measurement window is stated.
