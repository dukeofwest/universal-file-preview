# Live framework demo

The static demo combines the independent Angular, React, and Vue applications. The framework selector loads the chosen compiled app in an iframe; each app uses the published npm package and handles file selection, local preview, errors, clearing, and original-file download.

The repository is `dukeofwest/universal-file-preview`. The deployed Pages site is available at:

<https://dukeofwest.github.io/universal-file-preview/>

The public URL was opened and confirmed to serve the framework selector and compiled applications on 2026-11-02. The deployment workflow is [`deploy-pages.yml`](../.github/workflows/deploy-pages.yml). It builds and tests the library, builds all three applications, uploads the `site/` artifact, and deploys it to GitHub Pages. It runs after pushes to `main` and can also be started manually with `workflow_dispatch`.

## Enable GitHub Pages

If the repository does not already use GitHub Actions as its Pages source:

1. Open the repository on GitHub and select **Settings**.
2. Select **Pages** in the Code and automation section.
3. Under **Build and deployment**, choose **GitHub Actions** as the source.
4. Push the workflow and example changes to `main`, or use **Actions → Deploy framework demos to GitHub Pages → Run workflow** after the workflow exists on the default branch.
5. Wait for the build and deploy jobs to succeed. Open the published URL shown by the `github-pages` deployment environment and test the Angular, React, and Vue choices.

Users' chosen files are read in the browser by the library and are not uploaded by these examples. The site includes downloadable sample SVG, DOCX, XLSX, and PPTX files for trying the framework applications. These are synthetic demo fixtures. On 2026-11-02, all four samples were loaded successfully in each compiled Angular, React, and Vue app; this smoke check verifies the included samples, not compatibility with arbitrary Office documents or visual equivalence with Office.

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
