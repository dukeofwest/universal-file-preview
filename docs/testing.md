# Testing

Run the configured suite with:

```sh
npm test
npm run typecheck
```

`npm test` first runs the TypeScript build and then Node's built-in test runner over `tests/*.test.mjs`. The current suite has 20 tests covering binary reads, format detection, ZIP and XML handling, synthetic DOCX/XLSX/PPTX parser fixtures, and selected XLSX format/formula behavior.

The suite does not include browser-renderer integration, real-file visual comparisons, accessibility automation, or a supported-browser matrix. Passing it verifies only the tested code paths. `npm pack --dry-run` separately checks the files selected for publication; it is not a runtime consumer test.
