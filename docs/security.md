# Security and privacy

## Data flow

The library processes local `Blob`/`File` objects and byte arrays in the browser. It does not upload those inputs to a service. If a string URL is supplied, the browser fetches that URL; the application and remote server remain responsible for network access, authorization, and CORS policy.

## Untrusted documents

Treat every document as untrusted input. The implementation applies the limits and validation below. They reduce exposure, but they are not complete protection, and the gaps are listed after them.

- The viewer defaults to a 128 MiB source limit, configurable with `maxFileSize`.
- ZIP reading defaults to 32 MiB per entry and 4,096 entries, and rejects unsafe paths, duplicate paths, encryption, multi-disk archives, ZIP64, and unsupported compression methods.
- The XML parser defaults to 250,000 elements, nesting depth 256, and 16 MiB of source/text; DTDs and entity declarations are rejected, and there is no external entity resolution.
- External OOXML relationships are not followed to retrieve resources.
- OpenDocument spreadsheets have separate limits: `content.xml` or a whole `.fods` file 128 MiB, `styles.xml` 16 MiB, all parts together 256 MiB, 1,000,000 cells, 1,024 sheets, 32,767 characters per cell, and 10,000 elements per captured cell or style block. Exceeding one stops the preview with a `resource-limit` error rather than truncating content.
- Encrypted (password-protected) OpenDocument files are refused before any part is parsed. Macros are reported, not run; digital signatures are reported, not verified.
- OpenDocument external data (linked tables and DDE links) is never fetched, formulas are never evaluated, and only `http`, `https`, and `mailto` cell links are kept.
- Images are shown through `blob:` URLs, which have the origin of the page that creates them. All of them are created in one place:
  - Only PNG, JPEG, GIF, BMP, WebP, and SVG are accepted.
  - Raster images must start with a matching file signature, and the blob gets the type of the actual content, not the type suggested by a file name. Media whose content does not match is not shown.
  - SVG, both standalone `.svg` files and SVG media inside DOCX, PPTX, and XLSX packages, is reduced to an explicit allowlist first: only listed elements in the SVG namespace, only listed attributes without a namespace, each with a strict value rule (for example, paint values may refer only to an element inside the same image). Comments, processing instructions, and CDATA are removed. The result is parsed again and must pass the same allowlist unchanged, otherwise the image is not shown. This removes scripts, event handlers, styles, and external references, and intentionally removes many SVG features.
- Font names from documents are accepted only when they consist of letters, digits, spaces, and a few punctuation marks; any other name is replaced by a generic family. Styles are set one CSS property at a time, so a value from a document cannot add other declarations.
- The viewer root (`.ufp-root`) uses `contain: layout paint style`, `isolation: isolate`, and `overflow: hidden`, so document content cannot be positioned or painted outside the viewer, even with `position: fixed`. Keep these declarations if you restyle the viewer. This also applies to your own elements: anything placed inside `.ufp-root` with `position: fixed` is positioned relative to the viewer and clipped to it, so put overlays such as dialogs and menus outside the viewer root.
- The XLSX formula evaluator implements a limited subset in project code. It does not use `eval`, execute macros, or fetch linked workbooks.

### Known gaps

- **No total decompression budget for DOCX, PPTX, and XLSX.** Each ZIP entry is limited to 32 MiB, but there is no limit on the total amount decompressed, and the same entry can be decompressed many times. A small Office file can therefore make the page use a large amount of memory and CPU. (OpenDocument packages do have a 256 MiB total.)
- **No overall limit on XLSX computation.** Some workbook features (hyperlink ranges, formulas over large ranges, and others) can make a small file take a long time to process.
- **Parsing is synchronous** and runs on the page's main thread. While a large or hostile file is processed, the page does not respond, and closing or replacing the viewer cannot interrupt the parser.
- **Browser image decoding** runs in the browser process; this package cannot isolate the page from browser codec vulnerabilities or limit the memory a decoded image uses.

Until these gaps are closed, previewing untrusted uploads can freeze or crash the browser tab. Do not rely on this library alone to protect an application that previews files from untrusted users; use the deployment guidance below.

## Deployment for untrusted files

- **Restrict input** to the formats and sizes your application needs (`maxFileSize`, and checks before calling the viewer).
- **Isolate the viewer** when documents come from other users. The strongest option is to render previews on a separate origin (for example a dedicated preview domain, embedded with an `<iframe>`), so that the preview page shares no cookies, storage, or same-origin access with your main application. Only a separate *site* (a different registrable domain, not a subdomain) lets browsers with site isolation run the preview in its own process; otherwise a preview that freezes can still freeze the embedding page.
- **Apply a Content Security Policy.** The viewer needs:
  - `img-src blob: data:` (images are shown from `blob:` URLs; spreadsheet pattern fills use `data:` SVG images);
  - `style-src 'unsafe-inline'` (the viewer adds one `<style>` element to the page; inline styles set through the DOM, as the viewer does, are not restricted by CSP);
  - `connect-src` for the origins you load documents from with string URLs (not needed for `File`, `Blob`, or byte-array sources).

  A starting point for the page that hosts the viewer, to adapt to your application:

  ```
  Content-Security-Policy: default-src 'self'; script-src 'self'; img-src 'self' blob: data:; style-src 'self' 'unsafe-inline'; connect-src 'self'; font-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'self'
  ```

  Not allowing remote origins in `img-src` and `font-src` also prevents a document from making the browser contact a remote server through CSS, should such a bug exist.

  When a document contains SVG with event handler attributes, the browser may report blocked inline event handlers (`script-src-attr`) while the viewer sanitizes that SVG. The handlers never run; the reports are expected and show the policy adding a second layer of protection.
- **Trusted Types:** the viewer is not yet compatible with `require-trusted-types-for 'script'`. It sets a few static icons with `innerHTML` and parses SVG with `DOMParser`, which that policy blocks.

## Application responsibilities

- Apply your own access controls before allowing remote URLs or local documents to be previewed.
- Do not treat parsed document contents as trusted application HTML. Use the library renderer output as designed and avoid reinserting document strings with `innerHTML`.
- Surface `PreviewError` appropriately without exposing sensitive URLs or document content in public logs.
- Call `destroy()` when a viewer is removed so object URLs and active requests are cleaned up.
- Keep the package updated and review security notices for both this project and supported browsers.

## Reporting a vulnerability

The project is maintained in [dukeofwest/universal-file-preview](https://github.com/dukeofwest/universal-file-preview). Use [GitHub's private security advisory form](https://github.com/dukeofwest/universal-file-preview/security/advisories/new) when private vulnerability reporting is enabled. The setting has not been independently verified in this documentation update. If GitHub does not offer the private form, do not disclose the issue in a public tracker; first contact the repository maintainer privately to arrange a monitored secure channel. Do not attach confidential documents. See [SECURITY.md](../SECURITY.md) for the policy.
