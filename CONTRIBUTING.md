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
npm pack --dry-run
```

Add synthetic fixtures for successful and malformed input paths. Add appropriately licensed real-file fixtures and browser visual checks before making fidelity claims. Do not claim format support based only on detection, package recognition, or text extraction.

## Security and resource use

Never execute document-provided scripts or formulas, resolve external XML entities, or follow external package relationships. Keep bounds explicit and test new allocation/decompression paths. Do not commit credentials, private documents, or unlicensed fixtures.

## Pull requests

Describe user-visible behavior, limitations, and the commands you ran. Include screenshots only when they help explain a browser-rendering change and ensure sample documents can be shared legally. See [Security](./SECURITY.md) for private vulnerability reporting status and [Code of Conduct](./CODE_OF_CONDUCT.md) for community expectations.
