# Development and testing

## Prerequisites

Use Node.js 22, matching the current CI workflow. `npm ci` installs the lockfile's TypeScript development dependency. There are no runtime dependencies.

## Commands

```sh
npm ci
npm run build
npm run typecheck
npm test
npm pack --dry-run
npm run demo
```

- `npm run build` emits ESM JavaScript and declaration files into `dist/`.
- `npm run typecheck` checks the TypeScript source without emitting output.
- `npm test` builds and runs the Node.js test files with the built-in test runner.
- `npm pack --dry-run` lists the package contents without creating a tarball.
- `npm run demo` serves the demo at `http://127.0.0.1:4173/` after a build.

## Test scope

Tests are under `tests/`. They cover binary reads, format detection, ZIP validation/decompression, XML parsing, synthetic DOCX/XLSX/PPTX parser fixtures, and selected XLSX calculations/render-model features. They do not mount the viewer in a browser, compare rendered pages/slides against Office, or provide general compatibility evidence for arbitrary real documents.

When extending a parser, add positive and malformed synthetic fixtures for the behavior. Keep claims about rendered support aligned with tested parsing behavior, and use appropriately licensed real documents before making visual-fidelity claims. Do not add a third-party document parser or renderer.

## Build output and package contents

`dist/` is generated and ignored by `.gitignore`; it is included by the package allowlist in `package.json`. Run the build before packing. Tests, the demo, and development configuration are not included in the npm package. Review `npm pack --dry-run` whenever package metadata or the allowlist changes.
