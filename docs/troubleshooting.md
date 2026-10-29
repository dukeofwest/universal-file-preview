# Troubleshooting

## “Preview is not available”

The default registry renders GIF, PNG, JPEG, WebP, SVG, DOCX, XLSX, and PPTX. PDF and legacy `.doc`, `.xls`, and `.ppt` files can be detected but have no built-in renderer. Register an application-owned renderer for other formats. See [format support](./format-support.md).

## The file is rejected as invalid or unsupported

Check that the file is complete and not encrypted. OOXML packages must use supported ZIP methods and package structures. ZIP64, encrypted entries, multi-disk archives, unsupported compression, malformed XML, and unsafe package paths are rejected. A file extension or MIME type alone does not make an unsupported parser available.

## Remote file does not load

URL sources are fetched by the browser. Check the browser network panel, response status, authentication, and the remote server's CORS headers. A successful request in a server-side tool does not prove that the browser origin is allowed.

## “File exceeds the configured size limit”

The default source limit is 128 MiB. Set `maxFileSize` to a positive safe integer in bytes if the application needs a different limit. ZIP entries also have a built-in per-entry limit; increasing the source limit does not remove archive parser bounds.

## Missing or changed document styling

DOCX/XLSX/PPTX rendering implements subsets of the Office formats. Unsupported effects, formulas, charts, fonts, and layout rules can be omitted or approximated; see the format table for concrete gaps. Confirm the input format and compare against a supported feature before treating a visual difference as a regression. The current project has parser tests but no browser visual-regression suite.

## Missing styles in an application

The viewer injects the styles it needs under its own root. If an application uses the `ufp-*` classes outside that root, import `universal-file-preview/styles.css`. Confirm the package subpath is preserved by the application's bundler.

## TypeScript cannot resolve an import

The package exports ESM and generated declarations from `dist`. Build the package before using a local source checkout. The `styles.css` subpath is a stylesheet, not a TypeScript module. CommonJS output is not provided.

## Large or complex files are slow

Parsing runs synchronously on the main thread after input bytes are read, and rendering may create many DOM elements. Reduce file size or complexity, avoid repeatedly creating viewers without calling `destroy()`, and consider running preview in an isolated application context. The library does not currently use a worker.

## A file was detected but failed to render

Detection uses signatures where available and falls back to MIME type or extension. Detection is separate from parsing and rendering. Catch the rejected `open()` or `load()` promise; `PreviewError.code` can distinguish unsupported formats, malformed input, resource limits, and network failures.
