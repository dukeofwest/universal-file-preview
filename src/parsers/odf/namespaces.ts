import type { SpreadsheetFormulaDialect } from "../../model-extensions.js";

/**
 * OpenDocument namespace URIs (ODF 1.3 Part 3 §1.5; implementation specification §4.5).
 *
 * Internal module. ODF code compares these URIs and local names, never prefixed strings:
 * producers may use any prefix. ODF 1.0-1.4 share the ":1.0" core URIs.
 */
export const ODF = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  draw: "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
  number: "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
  meta: "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
  config: "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
  chart: "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
  manifest: "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
  xlink: "http://www.w3.org/1999/xlink",
  dc: "http://purl.org/dc/elements/1.1/",
} as const;

/** Formula-language namespaces seen in table:formula values (specification §5 and research §5.4). */
export const FORMULA_NAMESPACES: ReadonlyMap<string, SpreadsheetFormulaDialect> = new Map<string, SpreadsheetFormulaDialect>([
  ["urn:oasis:names:tc:opendocument:xmlns:of:1.2", "openformula"],
  // ODF 1.3 Part 3 §2.2.1 D.3 mentions of:1.3; accepted as well.
  ["urn:oasis:names:tc:opendocument:xmlns:of:1.3", "openformula"],
  ["http://openoffice.org/2004/formula", "ooo"],
  ["http://schemas.microsoft.com/office/excel/formula", "msoxl"],
]);

/** Producer extension namespaces: recognised, never required. To be confirmed against real files. */
export const EXT = {
  loext: "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
  calcext: "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
  tableooo: "http://openoffice.org/2009/table",
  officeooo: "http://openoffice.org/2009/office",
} as const;
