import { parseFods } from "./odf/flat.js";
import { parseOds } from "./odf/ods.js";
import { parseFodt } from "./odf/odt-flat.js";
import { parseOdt } from "./odf/odt.js";
import { parseDocx, parsePptx } from "./ooxml.js";
import { parseXlsx } from "./xlsx.js";
import type { DocumentModel, WorkerFormat } from "../worker/protocol.js";

/**
 * Parses a built-in document format into its model, without the DOM. Internal module, shared by the
 * renderers (main-thread mode) and the parser worker, so both modes run the same parsers with the same limits.
 */
export interface ParseModelOptions {
  /** Locale for text the XLSX parser formats itself (number formats). */
  readonly locale: string;
  /** Cooperative cancellation (OpenDocument parsers); the worker is terminated instead. */
  readonly signal?: AbortSignal;
}

export function parseModel(format: WorkerFormat, bytes: Uint8Array, name: string, options: ParseModelOptions): DocumentModel {
  const signal = options.signal ? { signal: options.signal } : {};
  switch (format) {
    case "docx": return parseDocx(bytes, name);
    case "pptx": return parsePptx(bytes, name);
    case "xlsx": return parseXlsx(bytes, name, { locale: options.locale });
    case "ods": case "ots": return parseOds(bytes, name, signal);
    case "fods": return parseFods(bytes, name, signal);
    case "odt": case "ott": return parseOdt(bytes, name, signal);
    case "fodt": return parseFodt(bytes, name, signal);
  }
}

/** The locale the viewer formats with: the browser's language, as the renderers have always used. */
export function viewerLocale(): string {
  return typeof navigator !== "undefined" && navigator.language ? navigator.language : "en-US";
}
