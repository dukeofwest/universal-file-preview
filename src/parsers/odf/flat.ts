import { DiagnosticCollector } from "../../diagnostics.js";
import { PreviewError } from "../../errors.js";
import type { ExtendedWorkbook } from "../../model-extensions.js";
import { decodeXmlText } from "../../package/package-reader.js";
import { resolveLimits, streamSpreadsheet, type OdsParseOptions } from "./ods.js";

/**
 * Flat OpenDocument spreadsheets (.fods; ODF 1.3 Part 3 §2.2.1, implementation specification
 * M15). Internal module.
 *
 * The whole document is one XML file whose office:document root names the spreadsheet media type
 * in office:mimetype. It is decoded with the package text rules and streamed once: font faces,
 * common and automatic styles, metadata and settings are captured, and office:body is read
 * exactly as content.xml of a package. Embedded images and objects are not kept.
 */
export function parseFods(bytes: Uint8Array, name: string, options: OdsParseOptions = {}): ExtendedWorkbook {
  const limits = resolveLimits(options.limits);
  if (bytes.length > limits.maxContentBytes) throw new PreviewError("The flat spreadsheet exceeds the configured size limit.", "resource-limit");
  const source = decodeXmlText(bytes, "The flat spreadsheet");
  return streamSpreadsheet({ source, part: "flat spreadsheet", mode: "flat", format: "fods", chartObjects: new Set() }, name, limits, new DiagnosticCollector(), options.signal);
}
