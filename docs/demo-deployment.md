# Demo deployment

The GitHub Pages site combines independent Angular, React, and Vue applications. The landing page selects and embeds the actual compiled framework application. Each app installs and imports the published `@dukeofwest/universal-file-preview@0.2.1` package. Selected files are read in the browser and are not uploaded by these examples.

## GitHub Pages

The GitHub Actions workflow at [`.github/workflows/deploy-pages.yml`](../.github/workflows/deploy-pages.yml) runs after pushes to `main` and can be started manually. It installs the core and example dependencies from their lockfiles, runs core tests and type checking, builds all three apps, uploads the static `site/` directory, and deploys it to GitHub Pages.

For `dukeofwest/universal-file-preview`, the live URL is <https://dukeofwest.github.io/universal-file-preview/>. It serves the framework selector and the compiled applications; see [Live framework demo](./live-demo.md) for the most recent verification. Future changes should be treated as deployed only after a successful workflow run is visible in GitHub Actions and the `github-pages` environment.

If Pages is not enabled for the repository:

1. Open **Settings → Pages** in the GitHub repository.
2. Set **Build and deployment → Source** to **GitHub Actions**.
3. Push the workflow to `main` and wait for its deployment job, or run the workflow manually from the Actions tab.

## Build and preview locally

Use Node.js 22.22.3 or newer on the 22.x line (or 24.15 or newer). Install the root dependencies and each framework app's dependencies, then build and serve the Pages output:

```sh
npm ci
npm ci --prefix examples/angular
npm ci --prefix examples/react
npm ci --prefix examples/vue
npm run build:pages
npm run preview:pages
```

Open <http://127.0.0.1:4174/>. The examples use relative asset URLs and Angular's base href is `./` so the apps can be hosted below the repository's Pages path.

## Existing package demo

The original lightweight demo remains available locally with `npm run demo` at <http://127.0.0.1:4173/>. It serves the library's direct browser example and is separate from the Pages application.
