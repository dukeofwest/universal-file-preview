import { DiagnosticCollector } from "../../diagnostics.js";
import { PreviewError } from "../../errors.js";
import type { ExtendedPresentation } from "../../model-extensions.js";
import { decodeXmlText } from "../../package/package-reader.js";
import { resolveOdpLimits, streamPresentation, type OdpParseOptions } from "./odp.js";

/**
 * Flat OpenDocument presentations (.fodp), one XML file holding what a package splits across parts
 * (ODF 1.3 Part 3 §2.2.1).
 *
 * Internal module: not exported from the package entry point (decision D2). The whole file is the
 * content stream, so the same reader handles it with `mode: "flat"`: the styles, master styles and
 * metadata arrive as captured subtrees instead of separate parts, and images are stored inline as
 * base64 rather than as package entries. There is no ZIP and therefore no decompression to bound;
 * the work a document may cost is bounded by the same element and object budgets.
 */
export function parseFodp(bytes: Uint8Array, name: string, options: OdpParseOptions = {}): ExtendedPresentation {
  const limits = resolveOdpLimits(options.limits);
  if (bytes.length > limits.maxContentBytes) {
    throw new PreviewError("The document exceeds the configured size limit.", "resource-limit");
  }
  const collector = new DiagnosticCollector();
  const source = decodeXmlText(bytes, "The document");
  return streamPresentation({ source, part: name, mode: "flat", format: "fodp" }, name, limits, collector, options.signal);
}
