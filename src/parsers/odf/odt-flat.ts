import { DiagnosticCollector } from "../../diagnostics.js";
import { PreviewError } from "../../errors.js";
import type { ExtendedWordDocument } from "../../model-extensions.js";
import { decodeXmlText } from "../../package/package-reader.js";
import { resolveOdtLimits, streamText, type OdtParseOptions } from "./odt.js";

/**
 * Flat OpenDocument text documents (.fodt), one XML file holding what a package splits across parts
 * (ODF 1.3 Part 3 §2.2.1).
 *
 * Internal module: not exported from the package entry point (decision D2). The whole file is the
 * content stream, so the same reader handles it with `mode: "flat"`: the styles, master styles and
 * metadata arrive as captured subtrees instead of separate parts, and images are stored inline as
 * base64 rather than as package entries. There is no ZIP and therefore no decompression to bound;
 * the work a document may cost is bounded by the same element and block budgets.
 */
export function parseFodt(bytes: Uint8Array, name: string, options: OdtParseOptions = {}): ExtendedWordDocument {
  const limits = resolveOdtLimits(options.limits);
  if (bytes.length > limits.maxContentBytes) {
    throw new PreviewError("The document exceeds the configured size limit.", "resource-limit");
  }
  const collector = new DiagnosticCollector();
  const source = decodeXmlText(bytes, "The document");
  return streamText({ source, part: name, mode: "flat", format: "fodt" }, name, limits, collector, options.signal);
}
