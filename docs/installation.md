# Installation

## Install from npm

Install the package with the package manager used by your application:

```sh
npm install @dukeofwest/universal-file-preview
```

```sh
pnpm add @dukeofwest/universal-file-preview
```

```sh
yarn add @dukeofwest/universal-file-preview
```

The commands use the package name configured in `package.json`.

## Requirements

- **Browser:** a DOM-capable browser with `Blob`, `URL.createObjectURL`, `fetch`, `AbortController`, and image element support; module workers for the optional worker mode. The project has no tested minimum browser-version matrix yet.
- **JavaScript:** the package is emitted as ES modules targeting ES2022. CommonJS output is not provided.
- **Node.js for development:** CI currently runs on Node.js 22. A minimum supported development Node.js version has not been declared.
- **Dependencies:** no runtime dependencies or peer dependencies. TypeScript is the only declared development dependency.
- **Styles:** the viewer injects its scoped styles. Import `@dukeofwest/universal-file-preview/styles.css` only if an application also needs the prefixed styles outside the viewer root.
- **Workers:** optional. Worker mode (the `worker` option) needs the package's `./worker` entry point started by your application as a worker; see [Set up worker mode](#set-up-worker-mode) and the [API reference](./api-reference.md#worker-mode). Without the option, no worker is used.
- **Remote sources:** the browser must be allowed to fetch the URL under the server's CORS policy.

## Set up worker mode

Worker mode parses documents in a Web Worker that your application creates. The package provides the worker's code as the `./worker` entry point; your bundler (or server) has to make it available as a worker script on your origin. The recipes below were checked with the packed package in the React, Vue (Vite), and Angular (`@angular/build:application`) example applications.

**Any bundler that supports `new Worker(new URL(…, import.meta.url))`** (for example Vite and the Angular application builder): add a one-line worker module to your application and point the worker at it:

```ts
// src/parser.worker.ts
import "@dukeofwest/universal-file-preview/worker";
```

```ts
const viewer = new FilePreviewer({
  container,
  file,
  worker: { create: () => new Worker(new URL("./parser.worker.ts", import.meta.url), { type: "module" }) },
});
```

Use the path of your worker module relative to the file that creates the viewer (Angular resolves `"./parser.worker"` without the extension). Bundlers do not accept a package name directly inside `new URL(…)`, which is why the local module is needed.

**Vite** can also import the entry point as a worker constructor (with `vite/client` types, for example in a `vite-env.d.ts` with `/// <reference types="vite/client" />`):

```ts
import ParserWorker from "@dukeofwest/universal-file-preview/worker?worker";

const viewer = new FilePreviewer({ container, file, worker: { create: () => new ParserWorker() } });
```

**Without a bundler**, serve the package's `dist` directory from your origin (the worker imports the modules next to it) and create a module worker for `dist/worker.js`:

```ts
worker: { create: () => new Worker("/vendor/universal-file-preview/dist/worker.js", { type: "module" }) }
```

The worker entry point runs only inside a worker; importing it in the page throws an error. Other bundlers (for example webpack) have not been tested. With a Content Security Policy, allow the worker's origin in `worker-src`; see [Security](./security.md#deployment-for-untrusted-files). Frameworks may need their own Content Security Policy settings (for example, Angular's inlined critical CSS and component styles); the viewer itself needs none of those.

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

The currently published version is `0.3.0`; after `npm pack`, install the generated `dukeofwest-universal-file-preview-0.3.0.tgz` from the project directory into the consumer. For a future version, use the filename printed by `npm pack`. This local workflow does not require the package to be published. `npm pack --dry-run` lists the files without creating a tarball.

## Bundlers and frameworks

The package exposes ESM at its root, a CSS subpath, and the optional `./worker` entry point. No framework-specific peer dependency is required; worker configuration is needed only for worker mode (see [Set up worker mode](#set-up-worker-mode)). Vite, Angular CLI, and other bundlers are not covered by this project's automated tests; the worker recipes above were checked by hand with the example applications. Use their normal ESM package resolution and verify the installed package in your application build.
