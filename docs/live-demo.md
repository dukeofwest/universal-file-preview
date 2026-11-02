# Live framework demo

The static demo combines the independent Angular, React, and Vue applications. The framework selector loads the chosen compiled app in an iframe; each app uses the published npm package and handles file selection, local preview, errors, clearing, and original-file download.

The repository is `dukeofwest/universal-file-preview`. After the Pages workflow completes successfully, the expected site URL is:

<https://dukeofwest.github.io/universal-file-preview/>

The public URL was checked on 2026-11-02 and currently returns GitHub Pages' 404 page. The demo is not deployed yet. The URL is not considered live until a successful deployment is visible in the repository's Actions and Pages settings. The deployment workflow is [`deploy-pages.yml`](../.github/workflows/deploy-pages.yml). It builds and tests the library, builds all three applications, uploads the `site/` artifact, and deploys it to GitHub Pages. It runs after pushes to `main` and can also be started manually with `workflow_dispatch`.

## Enable GitHub Pages

If the repository does not already use GitHub Actions as its Pages source:

1. Open the repository on GitHub and select **Settings**.
2. Select **Pages** in the Code and automation section.
3. Under **Build and deployment**, choose **GitHub Actions** as the source.
4. Push the workflow and example changes to `main`, or use **Actions → Deploy framework demos to GitHub Pages → Run workflow** after the workflow exists on the default branch.
5. Wait for the build and deploy jobs to succeed. Open the published URL shown by the `github-pages` deployment environment and test the Angular, React, and Vue choices.

Repository Pages settings and a completed deployment were not changed or verified as part of the local implementation. Users' chosen files are read in the browser by the library and are not uploaded by these examples.

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
