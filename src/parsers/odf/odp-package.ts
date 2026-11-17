import { PreviewError } from "../../errors.js";
import type { DiagnosticCollector } from "../../diagnostics.js";
import { inspectOdfPackage, type OdfPackageInfo } from "../../package/odf-package.js";
import type { PackageReader } from "../../package/package-reader.js";

/**
 * OpenDocument presentation package validation (ODF 1.3 Part 2), kept apart from the text and
 * spreadsheet parsers so that reading a presentation changes nothing about reading either of those.
 *
 * Internal module: not exported from the package entry point (decision D2). It checks the package
 * contract that `inspectOdfPackage` does not: that the declared media type is a presentation one,
 * and that the stream this library reads is present. Inspection has already refused encrypted
 * packages and reported signatures and macros without reading them.
 *
 * No compression-ratio limit is applied, for the reason recorded in ods.ts: OpenDocument XML
 * legitimately compresses far past the 100:1 ratio used for Office Open XML parts, and an attack can
 * be padded below any threshold that keeps real documents working. The work a document may cost is
 * bounded by its content instead, where the parser reads it.
 */

/** The presentation media types this library reads, and the format each is previewed as. */
const PRESENTATION_MEDIA_TYPES: Readonly<Record<string, "odp" | "otp">> = {
  "application/vnd.oasis.opendocument.presentation": "odp",
  "application/vnd.oasis.opendocument.presentation-template": "otp",
};

export interface OdpPackage {
  /** Decided by the package's own media type, never by the file name. */
  readonly format: "odp" | "otp";
  readonly info: OdfPackageInfo;
}

/**
 * Opens a presentation's package: the media type decides the format, and `content.xml` must exist.
 * Throws PreviewError("unsupported-format") for any other OpenDocument kind, including the drawing
 * and graphics types, and PreviewError("invalid-file") for a package without content.
 */
export function readOdpPackage(reader: PackageReader, collector: DiagnosticCollector): OdpPackage {
  const info = inspectOdfPackage(reader, collector);
  const format = Object.hasOwn(PRESENTATION_MEDIA_TYPES, info.mediaType) ? PRESENTATION_MEDIA_TYPES[info.mediaType] : undefined;
  if (!format) throw new PreviewError("Not an OpenDocument presentation.", "unsupported-format");
  if (!reader.has("content.xml")) throw new PreviewError("The presentation has no content.xml.", "invalid-file");
  return { format, info };
}
