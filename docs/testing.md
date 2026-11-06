# Testing

Run the configured suite with:

```sh
npm test
npm run typecheck
```

`npm test` first runs the TypeScript build and then Node's built-in test runner over `tests/*.test.mjs`. The current suite has 261 tests covering:

- binary reads, format detection (including OpenDocument detection), ZIP decoding and limits, and XML handling (the tree parser, the namespace-aware tokenizer, namespace resolution, and the streaming reader);
- characterization snapshots that pin the DOCX, XLSX, PPTX, and demo-sample parser output, plus synthetic DOCX/XLSX/PPTX parser fixtures and selected XLSX format/formula behaviour;
- OpenDocument package validation, value, style and text utilities, the ODS parser, a financial-accuracy suite (typed values, saved display text, currencies, formulas, hidden data, empty versus zero), number styles, view settings, flat `.fods` files, and the renderer registration;
- workbook layout helpers and parse diagnostics;
- security regressions: font-name validation, image type and signature checks, malicious SVG and font-name fixtures for DOCX, PPTX, XLSX, ODS, and FODS (`tests/helpers/security-fixtures.mjs`), source checks that renderers never build `cssText` or create image blobs outside `renderers/image-url.ts`, and the viewer containment rules;
- resource limits for hostile files (`tests/resource-limits.test.mjs`): repeated and unused package parts, the total decompression budget, the XML expansion limit, overlapping or inconsistent ZIP records (including data descriptors and entries that are never read), XLSX sheet, cell, hyperlink, conditional-format, filter-button, slicer, formula-work and formula-depth limits, XML attribute and name limits, TIFF tile validation, and image dimension limits.

All OpenDocument fixtures are synthetic and LibreOffice-shaped; no files from real producers are included yet.

## Browser regression (manual)

`scripts/browser-regression.mjs` renders the demo samples and synthetic OOXML fixtures in a local headless Edge or Chrome (through the DevTools protocol, with no added dependencies). It compares the DOM, stylesheet, and screenshot of each case, in en-US and de-DE, against a baseline:

```sh
node scripts/browser-regression.mjs record    # on the commit before a renderer change
npm run build
node scripts/browser-regression.mjs compare   # after the change
```

Baselines are written to the system temporary directory and are not committed. The script is a development aid; it is not part of `npm test` or CI.

## Browser security check (manual)

`scripts/browser-security.mjs` renders the malicious fixtures in the same way, with a local "attacker" server that records any request it receives from any tab:

```sh
npm run build
node scripts/browser-security.mjs   # exit code 1 if any check fails; --report <file.json> sets the report path
```

Positive controls first confirm that the harness detects an unsanitized SVG blob and an injected style declaration. It then checks that every image blob has an allowed type, that SVG blobs contain no script and run none when opened as a document, and that no element escapes the viewer. An SVG sanitizer table covers namespace confusion, CSS-escaped `url()` values, entities, and invalid roots: the output must be clean, keep legitimate content, load as an image, and run no script when opened. The script also checks that a sheet with thousands of overlapping merges and column ranges renders quickly and that sheet images release their previous URLs when the grid re-renders, the input-message tooltip at all four viewer edges (identical with and without containment), frozen panes, scrolling, and rendering under the Content Security Policy recommended in [Security](./security.md). Any uncaught page exception fails the check, each step has a timeout and the whole run has one, and attacker-server requests are counted again after every opened tab has closed. The browser version, environment, and every result are written to a JSON report (by default in the system temporary directory). Like the regression script, it needs a local Edge or Chrome and is not part of `npm test` or CI.

The suite does not include real-file visual comparisons, accessibility automation, or a supported-browser matrix. Passing it verifies only the tested code paths. `npm pack --dry-run` separately checks the files selected for publication; it is not a runtime consumer test.
