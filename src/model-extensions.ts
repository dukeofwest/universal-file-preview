import type { PreviewDiagnostic } from "./diagnostics.js";
import type { SpreadsheetCell, SpreadsheetColumn, SpreadsheetRange, SpreadsheetStyle, Workbook, WordDocument, Worksheet } from "./model.js";

/**
 * Format-neutral spreadsheet model extensions (implementation specification §4.4 and §6,
 * milestone M8).
 *
 * Internal module: not exported from the package entry point, and the public model types in
 * model.ts are unchanged (decision D2: internal spreadsheet models stay internal). Each
 * `Extended*` type adds optional fields to its public counterpart, so:
 * - every public `Workbook` (for example from parseXlsx) is a valid `ExtendedWorkbook`, and
 * - every `ExtendedWorkbook` can be used wherever a public `Workbook` is expected.
 * The shared workbook renderer can therefore accept both XLSX and OpenDocument models.
 *
 * Producers fill BOTH the legacy fields (value, type, display, width, rows) so existing
 * rendering keeps working, and these fields so information needed for faithful financial
 * display is not lost (specification §6.4).
 */

/** Where a model came from. Reused by the document and presentation models in later phases. */
export type DocumentSourceFormat = "xlsx" | "ods" | "ots" | "fods" | "odt" | "ott" | "fodt";
export type DocumentSource = {
  format: DocumentSourceFormat;
  /** Producing application as declared by the file, e.g. meta:generator. */
  producer?: string;
  /** Format version as declared by the file, e.g. office:version "1.3". */
  formatVersion?: string;
  /** BCP 47 tag of the document's default language, e.g. "de-DE". */
  locale?: string;
};

/**
 * A WordDocument with the fields the OpenDocument text parser adds. The public WordDocument is
 * unchanged (decision D2): renderers read these through this type, and the DOCX parser never sets them.
 */
export type ExtendedWordDocument = WordDocument & {
  source?: DocumentSource;
  diagnostics?: PreviewDiagnostic[];
};

export type SpreadsheetValueKind = "empty" | "number" | "percentage" | "currency" | "date" | "time" | "boolean" | "string" | "error" | "invalid";

/**
 * The typed value of a cell. `lexical` is the stored attribute text, verbatim; `number` is
 * derived from it and never from display text. Values that do not match their declared type
 * are `invalid` and are never coerced.
 */
export type SpreadsheetCellValue =
  | { kind: "empty" }
  /** A percentage is stored as its fraction: 0.15 for 15%. */
  | { kind: "number" | "percentage"; lexical: string; number: number }
  /** `currency` is the ISO 4217 code exactly as stored. */
  | { kind: "currency"; lexical: string; number: number; currency?: string }
  /** `serial` counts days from Workbook.calculation.nullDate; absent when it cannot be computed. */
  | { kind: "date"; lexical: string; serial?: number }
  /** A duration in days; absent when the duration has year or month components. */
  | { kind: "time"; lexical: string; days?: number }
  | { kind: "boolean"; lexical: string; boolean: boolean }
  | { kind: "string"; text: string }
  /** The saved error text, e.g. "#DIV/0!" or "Err:502". */
  | { kind: "error"; code: string }
  | { kind: "invalid"; declaredType: string; lexical?: string };

export type SpreadsheetFormulaDialect = "excel" | "openformula" | "ooo" | "msoxl" | "unknown";
export type SpreadsheetFormula = {
  /** The formula attribute verbatim, e.g. "of:=SUM([.A1:.A3])". */
  text: string;
  /** `text` without its resolved namespace prefix, e.g. "=SUM([.A1:.A3])". */
  expression: string;
  dialect: SpreadsheetFormulaDialect;
  /** Namespace URI of the resolved prefix, when it was bound. */
  namespace?: string;
  /** "cached": saved by the producer; "computed": calculated by this library; "missing": neither. */
  result: "cached" | "computed" | "missing";
  /** Range of a matrix (array) formula. */
  matrix?: SpreadsheetRange;
};

/** "file": text saved by the producing application; "computed": formatted by this library; "raw": the stored value unformatted. */
export type SpreadsheetDisplaySource = "file" | "computed" | "raw";
export type SpreadsheetDisplay = {
  /** What the grid shows; equals SpreadsheetCell.display when both are present. */
  text: string;
  source: SpreadsheetDisplaySource;
  /** This library's formatting of the value, set only when it differs from saved ("file") text. */
  computed?: string;
};

export type SpreadsheetCellData = {
  value: SpreadsheetCellValue;
  display: SpreadsheetDisplay;
  formula?: SpreadsheetFormula;
  /** The cell lies under a merged range. */
  covered?: boolean;
};

export type SpreadsheetNumberFormatCategory = "general" | "number" | "currency" | "percentage" | "scientific" | "fraction" | "date" | "time" | "boolean" | "text";
export type SpreadsheetNumberFormatInfo = {
  category: SpreadsheetNumberFormatCategory;
  /** Name of the source format definition, e.g. an ODF data style "N104". */
  sourceName?: string;
  /** Excel-style format code, as stored (XLSX) or derived (OpenDocument). */
  code?: string;
  /** BCP 47 tag of the format's language and country. */
  locale?: string;
  currencySymbol?: string;
  /** ISO 4217 code, when determinable. */
  currency?: string;
  decimalPlaces?: number;
  /** False when this library cannot reproduce the format; only saved text is then shown. */
  supported: boolean;
};

/** A run of rows with identical properties, e.g. from table:number-rows-repeated. Heights are in points. */
export type SpreadsheetRowRange = { min: number; max: number; height?: number; hidden?: boolean; filtered?: boolean; style?: number };
export type SpreadsheetDefinedName = { name: string; sheet?: string; ref?: string; expression?: string };

export type ExtendedSpreadsheetStyle = SpreadsheetStyle & { numberFormatInfo?: SpreadsheetNumberFormatInfo };
export type ExtendedSpreadsheetCell = SpreadsheetCell & { data?: SpreadsheetCellData };
/** `widthPoints` is an absolute width and takes precedence over the Excel-unit `width`. */
export type ExtendedSpreadsheetColumn = SpreadsheetColumn & { widthPoints?: number; filtered?: boolean };

/**
 * A worksheet with optional neutral fields. Renderers resolve a row's properties from
 * `rows[row]` first, then from the sorted, non-overlapping `rowRanges` (specification §4.4).
 */
export type ExtendedWorksheet = Omit<Worksheet, "cells" | "columns"> & {
  cells: ExtendedSpreadsheetCell[];
  columns?: ExtendedSpreadsheetColumn[];
  rowRanges?: SpreadsheetRowRange[];
  /** Absolute default column width; takes precedence over defaultColumnWidth. */
  defaultColumnWidthPoints?: number;
  protected?: boolean;
};

export type ExtendedWorkbook = Omit<Workbook, "sheets" | "styles"> & {
  sheets: ExtendedWorksheet[];
  styles?: ExtendedSpreadsheetStyle[];
  source?: DocumentSource;
  calculation?: {
    /** xsd:date epoch for date serials, e.g. "1899-12-30". */
    nullDate?: string;
    /** The producer calculated with displayed precision. */
    precisionAsShown?: boolean;
  };
  names?: SpreadsheetDefinedName[];
  diagnostics?: PreviewDiagnostic[];
};
