/**
 * Non-fatal parser diagnostics (implementation specification AD-6, §4.3).
 *
 * Internal module: not exported from the package entry point (decision D2). Parsers record
 * conditions that do not prevent a preview but affect how far it can be trusted, such as
 * cached formula results, hidden rows, or features that are not drawn. Fatal problems
 * still throw `PreviewError`.
 *
 * Messages are user-facing English text. They must never contain cell contents or other
 * document data, apart from short metadata parameters (producer name, format version),
 * which are length-limited.
 */

/**
 * - `accuracy`: the preview may differ from the author's application in ways that affect how
 *   values are interpreted. Always shown.
 * - `warning`: content is missing or approximated (for example, charts that are not drawn).
 * - `info`: metadata; not shown by default.
 */
export type PreviewDiagnosticSeverity = "info" | "warning" | "accuracy";

export interface PreviewDiagnostic {
  readonly code: string;
  readonly severity: PreviewDiagnosticSeverity;
  readonly message: string;
  /** Number of occurrences; at least 1. */
  readonly count: number;
  readonly sheet?: string;
  /** First occurrence, in A1 notation. */
  readonly cell?: string;
  /** Package part where the condition was found, e.g. "content.xml". */
  readonly part?: string;
}

export interface DiagnosticDetails {
  readonly sheet?: string;
  readonly cell?: string;
  readonly part?: string;
  /** Occurrences to record with this call; a positive safe integer. Defaults to 1. */
  readonly count?: number;
}

/** Short metadata substituted into catalogue messages. Values are trimmed to MAX_PARAMETER_LENGTH characters. */
export interface DiagnosticParams {
  readonly producer?: string;
  readonly version?: string;
}

interface DiagnosticDefinition {
  readonly severity: PreviewDiagnosticSeverity;
  /** Rendered when the list is read, so counts and plurals reflect every occurrence. */
  readonly message: (count: number, params: DiagnosticParams) => string;
}

const MAX_PARAMETER_LENGTH = 120;
const TRUNCATED_CODE = "diagnostics.truncated";
const codePattern = /^[a-z0-9]+(?:[.-][a-z0-9]+)*$/;
const severities: ReadonlySet<string> = new Set<PreviewDiagnosticSeverity>(["info", "warning", "accuracy"]);

const number = (count: number) => count.toLocaleString("en-US");
const plural = (count: number, one: string, many: string) => count === 1 ? one : many;
const fixed = (severity: PreviewDiagnosticSeverity, text: string): DiagnosticDefinition => ({ severity, message: () => text });

/** Catalogue of the diagnostic codes used in Phase 1 (implementation specification §4.3). */
export const DIAGNOSTIC_MESSAGES = {
  "odf.formula.cached": { severity: "accuracy", message: (_count, params) => `Formula results are shown as last saved${params.producer ? ` by ${params.producer}` : ""}; they were not recalculated.` },
  "odf.formula.missing-result": { severity: "accuracy", message: count => `${number(count)} formula ${plural(count, "cell has", "cells have")} no saved result and ${plural(count, "is", "are")} shown as "Cannot calculate".` },
  "odf.formula.unknown-dialect": { severity: "info", message: count => `${number(count)} ${plural(count, "formula uses", "formulas use")} an unrecognised formula language.` },
  "odf.display.raw": { severity: "accuracy", message: count => `${number(count)} ${plural(count, "value had", "values had")} no saved display text and ${plural(count, "is", "are")} shown unformatted.` },
  "odf.display.mismatch": { severity: "accuracy", message: count => `${number(count)} ${plural(count, "cell", "cells")}: preview formatting differs from the saved text; the saved text is shown.` },
  "odf.value.invalid": { severity: "accuracy", message: count => `${number(count)} ${plural(count, "cell contains a value that does", "cells contain values that do")} not match ${plural(count, "its", "their")} declared type.` },
  "odf.hidden.sheets": { severity: "accuracy", message: count => `${number(count)} hidden ${plural(count, "sheet is", "sheets are")} not shown as tabs.` },
  "odf.hidden.rows": { severity: "accuracy", message: count => `${number(count)} hidden ${plural(count, "row", "rows")} in the used range.` },
  "odf.hidden.columns": { severity: "accuracy", message: count => `${number(count)} hidden ${plural(count, "column", "columns")} in the used range.` },
  "odf.hidden.filtered-rows": { severity: "accuracy", message: count => `${number(count)} ${plural(count, "row is", "rows are")} hidden by a filter.` },
  "odf.precision-as-shown": fixed("accuracy", "This workbook calculates with displayed precision; saved results may differ from the full stored values."),
  "odf.covered-cell-content": { severity: "warning", message: count => `${number(count)} ${plural(count, "cell", "cells")} hidden by merged ranges ${plural(count, "contains a value", "contain values")}.` },
  "odf.bounds.clipped": fixed("warning", "Content beyond row 1,048,576 or column XFD is not shown."),
  "odf.unsupported.drawings": { severity: "warning", message: count => `${number(count)} ${plural(count, "image or shape is", "images or shapes are")} not shown.` },
  "odf.unsupported.charts": { severity: "warning", message: count => `${number(count)} ${plural(count, "chart is", "charts are")} not shown.` },
  "odf.unsupported.conditional-formatting": fixed("warning", "Conditional formatting is not applied."),
  "odf.unsupported.pivot": fixed("info", "Pivot tables are shown as their saved cell values."),
  "odf.external-links": fixed("warning", "This workbook references external data; saved values are shown and nothing is fetched."),
  "odf.package.no-manifest": fixed("info", "The package has no manifest."),
  "odf.package.mimetype-position": fixed("info", "The package's mimetype entry is not stored first or is compressed."),
  "odf.package.signed": fixed("info", "The document is digitally signed; the signature was not verified."),
  "odf.package.macros": fixed("info", "The document contains macros; they were not run."),
  "odf.meta.ignored": fixed("info", "Document metadata could not be read."),
  "odf.value.timezone": { severity: "info", message: count => `${number(count)} date ${plural(count, "value includes", "values include")} a time zone; ${plural(count, "it is", "they are")} shown with ${plural(count, "its", "their")} written date and time.` },
  "odf.style.cycle": { severity: "info", message: count => `${number(count)} style inheritance ${plural(count, "cycle was", "cycles were")} ignored.` },
  "odf.style.default-missing": fixed("info", "The document declares no default cell style; Liberation Sans 10 pt is assumed."),
  "odf.version.newer": { severity: "info", message: (_count, params) => `The document uses OpenDocument version ${params.version ?? "(unknown)"}; newer features may be missing.` },
  "odf.settings.ignored": fixed("info", "View settings could not be read."),
} as const satisfies Readonly<Record<string, DiagnosticDefinition>>;

export type DiagnosticCode = keyof typeof DIAGNOSTIC_MESSAGES;

type Entry = { code: string; severity: PreviewDiagnosticSeverity; render: (count: number) => string; count: number; sheet?: string; cell?: string; part?: string };

/**
 * Collects diagnostics during one parse. Entries are de-duplicated by code and sheet: repeated
 * reports add to the count and keep the first cell and part. Beyond `maxEntries` distinct
 * entries, further new entries are summarised by a single `diagnostics.truncated` info entry.
 */
export class DiagnosticCollector {
  readonly #maxEntries: number;
  readonly #entries = new Map<string, Entry>();
  #omitted = 0;

  constructor(maxEntries = 200) {
    if (!Number.isSafeInteger(maxEntries) || maxEntries < 1) throw new RangeError("maxEntries must be a positive safe integer.");
    this.#maxEntries = maxEntries;
  }

  /**
   * Records a diagnostic with an explicit severity and message. A `{n}` placeholder in the
   * message is replaced by the final occurrence count when the list is read.
   */
  add(code: string, severity: PreviewDiagnosticSeverity, message: string, details: DiagnosticDetails = {}): void {
    if (!codePattern.test(code) || code === TRUNCATED_CODE) throw new RangeError(`Invalid diagnostic code: ${code}`);
    if (!severities.has(severity)) throw new RangeError(`Invalid diagnostic severity: ${String(severity)}`);
    if (typeof message !== "string" || !message) throw new RangeError("A diagnostic message is required.");
    this.#record(code, severity, count => message.replaceAll("{n}", number(count)), details);
  }

  /** Records a catalogued diagnostic; severity and message come from DIAGNOSTIC_MESSAGES. */
  report(code: DiagnosticCode, details: DiagnosticDetails = {}, params: DiagnosticParams = {}): void {
    if (!Object.hasOwn(DIAGNOSTIC_MESSAGES, code)) throw new RangeError(`Unknown diagnostic code: ${String(code)}`);
    const definition: DiagnosticDefinition = DIAGNOSTIC_MESSAGES[code], safe = sanitizeParams(params);
    this.#record(code, definition.severity, count => definition.message(count, safe), details);
  }

  /** The recorded diagnostics in first-report order, as new objects on every call. */
  list(): PreviewDiagnostic[] {
    const output: PreviewDiagnostic[] = [...this.#entries.values()].map(entry => ({
      code: entry.code, severity: entry.severity, message: entry.render(entry.count), count: entry.count,
      ...(entry.sheet !== undefined ? { sheet: entry.sheet } : {}),
      ...(entry.cell !== undefined ? { cell: entry.cell } : {}),
      ...(entry.part !== undefined ? { part: entry.part } : {}),
    }));
    if (this.#omitted) output.push({ code: TRUNCATED_CODE, severity: "info", message: `${number(this.#omitted)} further ${plural(this.#omitted, "diagnostic was", "diagnostics were")} omitted.`, count: this.#omitted });
    return output;
  }

  #record(code: string, severity: PreviewDiagnosticSeverity, render: (count: number) => string, details: DiagnosticDetails): void {
    const amount = details.count ?? 1;
    if (!Number.isSafeInteger(amount) || amount < 1) throw new RangeError("Diagnostic count must be a positive safe integer.");
    const key = `${code}\u0000${details.sheet ?? ""}`, existing = this.#entries.get(key);
    if (existing) {
      existing.count = Math.min(Number.MAX_SAFE_INTEGER, existing.count + amount);
      if (existing.cell === undefined && details.cell !== undefined) existing.cell = details.cell;
      if (existing.part === undefined && details.part !== undefined) existing.part = details.part;
      return;
    }
    if (this.#entries.size >= this.#maxEntries) { this.#omitted = Math.min(Number.MAX_SAFE_INTEGER, this.#omitted + 1); return; }
    const entry: Entry = { code, severity, render, count: amount };
    if (details.sheet !== undefined) entry.sheet = details.sheet;
    if (details.cell !== undefined) entry.cell = details.cell;
    if (details.part !== undefined) entry.part = details.part;
    this.#entries.set(key, entry);
  }
}

function sanitizeParams(params: DiagnosticParams): DiagnosticParams {
  const clean = (value: string | undefined) => {
    if (value === undefined) return undefined;
    // Collapse control characters and whitespace runs; document metadata is untrusted.
    const text = value.replace(/[\u0000-\u001f\u007f\s]+/g, " ").trim();
    return text.length > MAX_PARAMETER_LENGTH ? `${text.slice(0, MAX_PARAMETER_LENGTH - 1)}…` : text;
  };
  const producer = clean(params.producer), version = clean(params.version);
  return { ...(producer ? { producer } : {}), ...(version ? { version } : {}) };
}
