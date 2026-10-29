# Installation

## Install from npm

Install the package with the package manager used by your application:

```sh
npm install universal-file-preview
```

```sh
pnpm add universal-file-preview
```

```sh
yarn add universal-file-preview
```

The commands use the package name configured in `package.json`.

## Requirements

- **Browser:** a DOM-capable browser with `Blob`, `URL.createObjectURL`, `fetch`, `AbortController`, `DOMParser`, and image element support. The project has no tested minimum browser-version matrix yet.
- **JavaScript:** the package is emitted as ES modules targeting ES2022. CommonJS output is not provided.
- **Node.js for development:** CI currently runs on Node.js 22. A minimum supported development Node.js version has not been declared.
- **Dependencies:** no runtime dependencies or peer dependencies. TypeScript is the only declared development dependency.
- **Styles:** the viewer injects its scoped styles. Import `universal-file-preview/styles.css` only if an application also needs the prefixed styles outside the viewer root.
- **Workers:** none are required or emitted.
- **Remote sources:** the browser must be allowed to fetch the URL under the server's CORS policy.

## Use a local checkout

To work on the source, use the repository directory you received and run:

```sh
npm ci
npm run build
npm test
npm run typecheck
```

To consume a local packed build from another application, build first and then pack/install the tarball locally:

```sh
npm run build
npm pack
```

Install the resulting `universal-file-preview-0.1.0.tgz` from the project directory into the consumer. This local workflow does not require the package to be published. `npm pack --dry-run` lists the files without creating a tarball.

## Bundlers and frameworks

The package exposes ESM at its root and a CSS subpath. No framework-specific peer dependency or worker configuration is required. Vite, Angular CLI, and other bundlers are not separately covered by this project's automated tests; use their normal ESM package resolution and verify the installed package in your application build.
