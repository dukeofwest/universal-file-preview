# Live framework demo

The static demo combines the independent Angular, React, and Vue applications. The framework selector loads the chosen compiled app in an iframe; each app uses the published npm package and handles file selection, local preview, errors, clearing, and original-file download.

The repository is `dukeofwest/universal-file-preview`. The deployed Pages site is available at:

<https://dukeofwest.github.io/universal-file-preview/>

The site is deployed from `main` and serves the framework selector and the compiled applications. The deployment workflow is [`deploy-pages.yml`](../.github/workflows/deploy-pages.yml). It builds and tests the library, builds all three applications, uploads the `site/` artifact, and deploys it to GitHub Pages. It runs after pushes to `main` and can also be started manually with `workflow_dispatch`.

## Enable GitHub Pages

If the repository does not already use GitHub Actions as its Pages source:

1. Open the repository on GitHub and select **Settings**.
2. Select **Pages** in the Code and automation section.
3. Under **Build and deployment**, choose **GitHub Actions** as the source.
4. Push the workflow and example changes to `main`, or use **Actions → Deploy framework demos to GitHub Pages → Run workflow** after the workflow exists on the default branch.
5. Wait for the build and deploy jobs to succeed. Open the published URL shown by the `github-pages` deployment environment and test the Angular, React, and Vue choices.

Users' chosen files are read in the browser by the library and are not uploaded by these examples. The site includes downloadable sample SVG, DOCX, XLSX, PPTX, ODS, and FODS files for trying the framework applications. These are synthetic demo fixtures.

On 2026-11-04 an automated headless-Edge check of the deployed site (run from a maintainer machine through the Chrome DevTools Protocol; not part of CI) confirmed that the framework selector loads each app, and that the Angular, React, and Vue apps each previewed all six samples, showed an error for an unsupported file and for two damaged Office files, recovered when another file was chosen, downloaded a byte-identical original, removed the viewer and its injected style on **Clear**, and loaded all assets from below the site path without console errors. This verifies the included samples in one Chromium-based browser; it is not a cross-browser test, a visual comparison with Office or LibreOffice, or evidence of compatibility with arbitrary documents.

## Build and preview locally

From the repository root, install each standalone application and build the site:

```sh
npm ci
npm ci --prefix examples/angular
npm ci --prefix examples/react
npm ci --prefix examples/vue
npm run build:pages
npm run preview:pages
```

Open <http://127.0.0.1:4174/>. The app builds use relative asset paths and the Angular `baseHref` is `./`, so each framework application works under the repository subpaths as well as local preview.
