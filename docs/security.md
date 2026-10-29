# Security and privacy

## Data flow

The library processes local `Blob`/`File` objects and byte arrays in the browser. It does not upload those inputs to a service. If a string URL is supplied, the browser fetches that URL; the application and remote server remain responsible for network access, authorization, and CORS policy.

## Untrusted documents

Treat every document as untrusted input. The implementation applies explicit limits and validation:

- The viewer defaults to a 128 MiB source limit, configurable with `maxFileSize`.
- ZIP reading defaults to 32 MiB per entry and 4,096 entries, and rejects unsafe paths, duplicate paths, encryption, multi-disk archives, ZIP64, and unsupported compression methods.
- The XML parser defaults to 250,000 elements, nesting depth 256, and 16 MiB of source/text; DTDs and entity declarations are rejected, and there is no external entity resolution.
- External OOXML relationships are not followed to retrieve resources.
- SVG is reduced to an explicit element and attribute allowlist before it is displayed as an image. This intentionally removes SVG features.
- The XLSX formula evaluator implements a limited subset in project code. It does not use `eval`, execute macros, or fetch linked workbooks.

These controls reduce exposure but do not make parsing risk-free. Browser-native image decoding remains part of the browser process. Parsing is synchronous and can consume CPU and memory; large files can affect the page. Consider restricting allowed formats and file size in applications that accept user uploads.

## Application responsibilities

- Apply your own access controls before allowing remote URLs or local documents to be previewed.
- Do not treat parsed document contents as trusted application HTML. Use the library renderer output as designed and avoid reinserting document strings with `innerHTML`.
- Surface `PreviewError` appropriately without exposing sensitive URLs or document content in public logs.
- Call `destroy()` when a viewer is removed so object URLs and active requests are cleaned up.
- Keep the package updated and review security notices for both this project and supported browsers.

## Reporting a vulnerability

The project is maintained in [dukeofwest/universal-file-preview](https://github.com/dukeofwest/universal-file-preview). Use [GitHub's private security advisory form](https://github.com/dukeofwest/universal-file-preview/security/advisories/new) when private vulnerability reporting is enabled. The setting has not been independently verified in this documentation update. If GitHub does not offer the private form, do not disclose the issue in a public tracker; first contact the repository maintainer privately to arrange a monitored secure channel. Do not attach confidential documents. See [SECURITY.md](../SECURITY.md) for the policy.
