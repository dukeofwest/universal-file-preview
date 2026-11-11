import type { PreviewDiagnostic } from "../diagnostics.js";
import type { SpreadsheetRow } from "../model.js";
import type { ExtendedSpreadsheetCell, ExtendedSpreadsheetColumn, ExtendedSpreadsheetStyle, ExtendedWorkbook, ExtendedWorksheet, SpreadsheetRowRange } from "../model-extensions.js";

/**
 * Pure layout and precedence rules for the workbook grid (implementation specification §4.4,
 * §4.6 and milestone M9b). DOM-free, so they are unit-tested in Node. Internal module.
 *
 * Every fallback branch is the grid renderer's previous code, unchanged: models without the
 * optional extension fields (every XLSX model) lay out exactly as before.
 */

/** Excel column width in characters to pixels, via the default font's maximum digit width. */
export function columnPixels(width: number, mdw: number): number { return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw); }

/** Unzoomed pixel width of a column: hidden is 0, then an absolute width in points, then the Excel width, then the sheet default. */
export function columnWidthPixels(column: ExtendedSpreadsheetColumn | undefined, defaultPixels: number, mdw: number): number {
  if (column?.hidden) return 0;
  if (column?.widthPoints !== undefined) return column.widthPoints * 4 / 3;
  return column?.width !== undefined ? columnPixels(column.width, mdw) : defaultPixels;
}

/** Unzoomed default column width; a negative stored width marks an Excel base width (padding added). */
export function defaultColumnPixels(sheet: ExtendedWorksheet, mdw: number): number {
  if (sheet.defaultColumnWidthPoints !== undefined) return sheet.defaultColumnWidthPoints * 4 / 3;
  const stored = sheet.defaultColumnWidth;
  if (stored !== undefined && stored > 0) return columnPixels(stored, mdw);
  const base = stored !== undefined && stored < 0 ? -stored : 8;
  return Math.ceil((base * mdw + 5) / 8) * 8;
}

/**
 * Properties of a 1-based row: `rows[row]` first, then the run in `rowRanges` containing it.
 * `rowRanges` must be sorted by `min` and must not overlap (specification §4.4).
 */
export function rowInfoAt(sheet: ExtendedWorksheet, row: number): SpreadsheetRow | undefined {
  const direct = sheet.rows?.[row];
  if (direct !== undefined || !sheet.rowRanges?.length) return direct;
  return findRowRange(sheet.rowRanges, row);
}

function findRowRange(ranges: readonly SpreadsheetRowRange[], row: number): SpreadsheetRowRange | undefined {
  let low = 0, high = ranges.length - 1;
  while (low <= high) {
    const middle = (low + high) >>> 1, range = ranges[middle]!;
    if (row < range.min) high = middle - 1;
    else if (row > range.max) low = middle + 1;
    else return range;
  }
  return undefined;
}

/**
 * Whether a number that does not fit its column may be shortened with General formatting
 * (Excel behaviour) instead of showing "####". Never for values with saved display text or an
 * explicit number format: they must not be re-formatted with the viewer's locale
 * (financial invariants I-3 and I-11).
 */
export function allowsGeneralShrink(cell: ExtendedSpreadsheetCell, style: ExtendedSpreadsheetStyle | undefined): boolean {
  if (style?.numberFormat) return false;
  if (style?.numberFormatInfo && style.numberFormatInfo.category !== "general") return false;
  return !cell.data || cell.data.display.source === "computed";
}

/** Diagnostics for the notice above the grid: accuracy first, then warnings; info is not shown. */
export function noticeDiagnostics(workbook: ExtendedWorkbook): PreviewDiagnostic[] {
  return noticeEntries(workbook.diagnostics);
}

/** The same selection for any model that carries diagnostics (text documents as well as workbooks). */
export function noticeEntries(diagnostics: readonly PreviewDiagnostic[] | undefined): PreviewDiagnostic[] {
  const entries = diagnostics ?? [];
  return [...entries.filter(entry => entry.severity === "accuracy"), ...entries.filter(entry => entry.severity === "warning")];
}

/** Accessible name of a grid cell: reference and shown text, plus the currency and saved-result facts sighted users cannot see. */
export function cellAccessibleLabel(ref: string, text: string, cell: ExtendedSpreadsheetCell): string {
  let label = `${ref}: ${text}`;
  const value = cell.data?.value;
  if (value?.kind === "currency" && value.currency) label += `, currency ${value.currency}`;
  if (cell.data?.formula?.result === "cached") label += ", saved formula result";
  return label;
}
