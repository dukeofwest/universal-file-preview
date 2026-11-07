# Contributing

Contributions should preserve the existing architecture and its core constraint: document parsing and rendering remain implemented by this project. Do not add, install, or integrate third-party document parsing or rendering libraries.

## Before changing code

- Keep each format parser separate and pass typed document models to renderers.
- Check the current support matrix before describing a feature as supported.
- Explain any proposed runtime dependency and why browser APIs or existing project code cannot meet the need. A document parser or renderer dependency is not acceptable.
- Keep examples and documentation aligned with the actual public exports and behavior.

## Validation

Run the configured checks before submitting a change:

```sh
npm ci
npm test
npm run typecheck
npm run test:security     # denial-of-service corpus, run alone
npm run test:performance  # throughput benchmarks, run alone
npm pack --dry-run
```

The last two measure time and memory, so they run one file at a time; do not move timing limits into the parallel `npm test` run. For changes to parsers, also run the fuzzer for a few minutes (`node scripts/fuzz.mjs --seconds 300`) and, with a local Edge or Chrome, `node scripts/browser-security.mjs`. CI runs all of these.

Add synthetic fixtures for successful and malformed input paths. Add appropriately licensed real-file fixtures and browser visual checks before making fidelity claims. Do not claim format support based only on detection, package recognition, or text extraction.

## Security and resource use

Never execute document-provided scripts or formulas, resolve external XML entities, or follow external package relationships. Keep bounds explicit and test new allocation/decompression paths. Do not commit credentials, private documents, or unlicensed fixtures.

## Pull requests

Describe user-visible behavior, limitations, and the commands you ran. Include screenshots only when they help explain a browser-rendering change and ensure sample documents can be shared legally. See [Security](./SECURITY.md) for private vulnerability reporting status and [Code of Conduct](./CODE_OF_CONDUCT.md) for community expectations.
