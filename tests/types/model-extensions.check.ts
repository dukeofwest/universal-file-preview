// Compile-only checks for src/model-extensions.ts (run by tests/model-extensions.test.mjs).
// Nothing here executes; tsc must accept every line and every @ts-expect-error must be needed.
import type { Workbook, Worksheet, SpreadsheetCell } from "../../src/model.js";
import type { ExtendedWorkbook, ExtendedWorksheet, ExtendedSpreadsheetCell, SpreadsheetCellValue, SpreadsheetFormula, SpreadsheetDisplay, SpreadsheetNumberFormatInfo } from "../../src/model-extensions.js";

declare const publicWorkbook: Workbook;
declare const extendedWorkbook: ExtendedWorkbook;
declare const publicSheet: Worksheet;
declare const extendedSheet: ExtendedWorksheet;
declare const publicCell: SpreadsheetCell;

// Every public model is a valid extended model, and the reverse.
export const acceptsPublicWorkbook: ExtendedWorkbook = publicWorkbook;
export const usableAsPublicWorkbook: Workbook = extendedWorkbook;
export const acceptsPublicSheet: ExtendedWorksheet = publicSheet;
export const usableAsPublicSheet: Worksheet = extendedSheet;
export const acceptsPublicCell: ExtendedSpreadsheetCell = publicCell;

// An ODS-style currency cell with saved display text and a cached formula result.
export const currencyCell: ExtendedSpreadsheetCell = {
  ref: "B2", value: "1234.5", type: "n", display: "1.234,50 EUR", style: 3, formula: "=SUM([.B1:.B1])",
  data: {
    value: { kind: "currency", lexical: "1234.5", number: 1234.5, currency: "EUR" },
    display: { text: "1.234,50 EUR", source: "file" },
    formula: { text: "of:=SUM([.B1:.B1])", expression: "=SUM([.B1:.B1])", dialect: "openformula", namespace: "urn:oasis:names:tc:opendocument:xmlns:of:1.2", result: "cached" },
  },
};

export const workbook: ExtendedWorkbook = {
  kind: "workbook", title: "report.ods",
  sheets: [{
    name: "Q1", cells: [currencyCell], maxRow: 2, maxColumn: 2,
    columns: [{ min: 1, max: 3, widthPoints: 64.01 }],
    rowRanges: [{ min: 3, max: 1_048_576, height: 12.8, hidden: true }],
    defaultColumnWidthPoints: 64.01,
  }],
  styles: [{ numberFormatInfo: { category: "currency", sourceName: "N104", locale: "de-DE", currencySymbol: "EUR", currency: "EUR", decimalPlaces: 2, supported: true } }],
  source: { format: "ods", producer: "LibreOffice/24.8", formatVersion: "1.3", locale: "de-DE" },
  calculation: { nullDate: "1899-12-30", precisionAsShown: false },
  names: [{ name: "Total", sheet: "Q1", ref: "$Q1.$B$2" }],
  diagnostics: [{ code: "odf.formula.cached", severity: "accuracy", message: "Formula results are shown as last saved; they were not recalculated.", count: 1 }],
};

export const values: SpreadsheetCellValue[] = [
  { kind: "empty" },
  { kind: "number", lexical: "0", number: 0 },
  { kind: "percentage", lexical: "0.15", number: 0.15 },
  { kind: "date", lexical: "2026-10-28", serial: 46323 },
  { kind: "date", lexical: "2026-10-28T10:00:00+02:00" },
  { kind: "time", lexical: "PT36H", days: 1.5 },
  { kind: "time", lexical: "P1M" },
  { kind: "boolean", lexical: "true", boolean: true },
  { kind: "string", text: "1,000" },
  { kind: "error", code: "#DIV/0!" },
  { kind: "invalid", declaredType: "float", lexical: "abc" },
];

// Invalid shapes are rejected.
// @ts-expect-error a number value needs its derived number
export const missingNumber: SpreadsheetCellValue = { kind: "number", lexical: "1" };
// @ts-expect-error unknown value kinds are not allowed
export const unknownKind: SpreadsheetCellValue = { kind: "money", lexical: "1" };
// @ts-expect-error a string value has text, not lexical
export const stringWithLexical: SpreadsheetCellValue = { kind: "string", lexical: "x" };
// @ts-expect-error optional fields cannot be set to undefined (exactOptionalPropertyTypes)
export const explicitUndefined: SpreadsheetFormula = { text: "x", expression: "x", dialect: "unknown", result: "missing", namespace: undefined };
// @ts-expect-error unknown display sources are not allowed
export const badSource: SpreadsheetDisplay = { text: "x", source: "guessed" };
// @ts-expect-error number format info must say whether it is supported
export const missingSupported: SpreadsheetNumberFormatInfo = { category: "number" };
// @ts-expect-error formula results are cached, computed or missing
export const badResult: SpreadsheetFormula = { text: "x", expression: "x", dialect: "excel", result: "stale" };
