# Testing

Run the configured suite with:

```sh
npm test
npm run typecheck
```

`npm test` first runs the TypeScript build and then Node's built-in test runner over `tests/*.test.mjs`. The current suite has 229 tests covering:

- binary reads, format detection (including OpenDocument detection), ZIP decoding and limits, and XML handling (the tree parser, the namespace-aware tokenizer, namespace resolution, and the streaming reader);
- characterization snapshots that pin the DOCX, XLSX, PPTX, and demo-sample parser output, plus synthetic DOCX/XLSX/PPTX parser fixtures and selected XLSX format/formula behaviour;
- OpenDocument package validation, value, style and text utilities, the ODS parser, a financial-accuracy suite (typed values, saved display text, currencies, formulas, hidden data, empty versus zero), number styles, view settings, flat `.fods` files, and the renderer registration;
- workbook layout helpers and parse diagnostics.

All OpenDocument fixtures are synthetic and LibreOffice-shaped; no files from real producers are included yet.

## Browser regression (manual)

`scripts/browser-regression.mjs` renders the demo samples and synthetic OOXML fixtures in a local headless Edge or Chrome (through the DevTools protocol, with no added dependencies). It compares the DOM, stylesheet, and screenshot of each case, in en-US and de-DE, against a baseline:

```sh
node scripts/browser-regression.mjs record    # on the commit before a renderer change
npm run build
node scripts/browser-regression.mjs compare   # after the change
```

Baselines are written to the system temporary directory and are not committed. The script is a development aid; it is not part of `npm test` or CI.

The suite does not include real-file visual comparisons, accessibility automation, or a supported-browser matrix. Passing it verifies only the tested code paths. `npm pack --dry-run` separately checks the files selected for publication; it is not a runtime consumer test.
