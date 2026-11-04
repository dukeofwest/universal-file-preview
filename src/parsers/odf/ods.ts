import { DiagnosticCollector } from "../../diagnostics.js";
import { PreviewError } from "../../errors.js";
import type { SpreadsheetRange } from "../../model.js";
import type { ExtendedSpreadsheetCell, ExtendedSpreadsheetColumn, ExtendedWorkbook, ExtendedWorksheet, SpreadsheetCellData, SpreadsheetCellValue, SpreadsheetDefinedName, SpreadsheetDisplay, SpreadsheetFormula, SpreadsheetRowRange } from "../../model-extensions.js";
import { inspectOdfPackage } from "../../package/odf-package.js";
import { PackageReader } from "../../package/package-reader.js";
import type { ExpandedName } from "../../xml/xml-namespaces.js";
import { nsAttr, nsChildren, nsFirst, nsValuePrefix, parseXmlNs, type NsXmlNode } from "../../xml/ns-tree.js";
import { streamXml, type StartAction, type XmlStartElement } from "../../xml/xml-stream.js";
import { columnName } from "../xlsx.js";
import { EXT, FORMULA_NAMESPACES, ODF } from "./namespaces.js";
import { createStyleResolver, type OdfStyleResolver } from "./styles.js";
import { annotationText, cellHyperlink, cellRuns, savedCellText } from "./text.js";
import { DEFAULT_NULL_DATE, readCellValue } from "./values.js";

/**
 * OpenDocument spreadsheet parser (implementation specification §4.6, §6 and milestone M11).
 *
 * Internal module: not exported from the package entry point (decision D2). It produces the
 * format-neutral ExtendedWorkbook rendered by renderWorkbook. content.xml is streamed; only one
 * cell (and the style blocks) are held as trees at a time. Values come from value attributes;
 * the producer's saved text is what is displayed; formulas are never recalculated.
 */

export type OdsLimits = {
  /** Decoded size of content.xml. Default 128 MiB. */
  maxContentBytes: number;
  /** Decoded size of styles.xml. Default 16 MiB. */
  maxStylesBytes: number;
  /** All decompressed parts together. Default 256 MiB. */
  maxTotalBytes: number;
  /** Materialised cells across the workbook (content cells and clipped styled empty cells). Default 1,000,000. */
  maxCells: number;
  /** Default 1,024. */
  maxSheets: number;
  /** Characters of saved text in one cell. Default 32,767. */
  maxCellTextLength: number;
  /** Elements in one captured cell or style block. Default 10,000. */
  maxCaptureElements: number;
};
export type OdsParseOptions = { readonly signal?: AbortSignal; readonly limits?: Partial<OdsLimits> };

const DEFAULT_LIMITS: OdsLimits = { maxContentBytes: 128 * 1024 * 1024, maxStylesBytes: 16 * 1024 * 1024, maxTotalBytes: 256 * 1024 * 1024, maxCells: 1_000_000, maxSheets: 1024, maxCellTextLength: 32_767, maxCaptureElements: 10_000 };
const ODS_MEDIA_TYPES: Readonly<Record<string, "ods" | "ots">> = { "application/vnd.oasis.opendocument.spreadsheet": "ods", "application/vnd.oasis.opendocument.spreadsheet-template": "ots" };
const CHART_MEDIA_TYPE = "application/vnd.oasis.opendocument.chart";
const MAX_ROW = 1_048_576, MAX_COLUMN = 16_384;
/** LibreOffice's defaults when a sheet declares no column width or row height (layout only). */
const DEFAULT_COLUMN_WIDTH_POINTS = 2.258 * 72 / 2.54, DEFAULT_ROW_HEIGHT_POINTS = 0.178 * 72;
/** Styled empty cells are materialised only inside the grid the renderer draws (xlsx-renderer rowCount/columnCount). */
const RENDER_MARGIN_ROWS = 40, RENDER_MARGIN_COLUMNS = 12, RENDER_MIN_ROWS = 60, RENDER_MIN_COLUMNS = 26;

export function parseOds(bytes: Uint8Array, name: string, options: OdsParseOptions = {}): ExtendedWorkbook {
  const limits = resolveLimits(options.limits);
  const collector = new DiagnosticCollector();
  const reader = PackageReader.openZip(bytes, { partBytes: { "content.xml": limits.maxContentBytes, "styles.xml": limits.maxStylesBytes }, maxTotalBytes: limits.maxTotalBytes });
  const info = inspectOdfPackage(reader, collector);
  const format = ODS_MEDIA_TYPES[info.mediaType];
  if (!Object.hasOwn(ODS_MEDIA_TYPES, info.mediaType) || !format) throw new PreviewError("Not an OpenDocument spreadsheet.", "unsupported-format");

  const stylesRoot = reader.has("styles.xml") ? parsePart(() => parseXmlNs(reader.readText("styles.xml"), { maxSourceLength: limits.maxStylesBytes, ...(options.signal ? { signal: options.signal } : {}) }), "styles.xml") : undefined;
  if (stylesRoot && (stylesRoot.uri !== ODF.office || stylesRoot.local !== "document-styles")) throw new PreviewError("styles.xml is not an OpenDocument styles stream.", "invalid-file");
  if (!reader.has("content.xml")) throw new PreviewError("The spreadsheet has no content.xml.", "invalid-file");
  const chartObjects = new Set([...info.manifest.values()].filter(entry => entry.mediaType === CHART_MEDIA_TYPE).map(entry => entry.path.replace(/\/$/, "")));
  const producer = info.generator?.split(/[\s$]/)[0] || undefined;

  const content = new ContentReader(stylesRoot, collector, limits, chartObjects);
  const source = reader.readText("content.xml");
  parsePart(() => streamXml(source, content, { maxSourceLength: limits.maxContentBytes, maxElements: Math.max(250_000, Math.ceil(limits.maxContentBytes / 4)), maxTokens: Math.max(5_000_000, limits.maxContentBytes), maxCaptureElements: limits.maxCaptureElements, ...(options.signal ? { signal: options.signal } : {}) }), "content.xml");

  const sheets = content.sheets;
  if (!sheets.length) throw new PreviewError("ODS workbook contains no readable sheets.", "invalid-file");
  const hiddenSheets = sheets.filter(sheet => sheet.hidden).length;
  if (hiddenSheets) collector.report("odf.hidden.sheets", { count: hiddenSheets });
  if (content.formulaCells) collector.report("odf.formula.cached", { count: content.formulaCells }, producer ? { producer } : {});
  if (content.missingResults) collector.report("odf.formula.missing-result", { count: content.missingResults });
  if (content.precisionAsShown) collector.report("odf.precision-as-shown");

  const resolver = content.resolver();
  const locale = resolver.locale;
  return {
    kind: "workbook", title: name, sheets,
    styles: resolver.styles, defaultFont: resolver.defaultFont,
    source: { format, ...(producer ? { producer } : {}), ...(content.formatVersion ? { formatVersion: content.formatVersion } : {}), ...(locale ? { locale } : {}) },
    calculation: { nullDate: content.nullDate, ...(content.precisionAsShown ? { precisionAsShown: true } : {}) },
    ...(content.names.length ? { names: content.names } : {}),
    diagnostics: collector.list(),
  };
}

function resolveLimits(overrides: Partial<OdsLimits> | undefined): OdsLimits {
  const limits = { ...DEFAULT_LIMITS };
  for (const [key, value] of Object.entries(overrides ?? {})) {
    if (!Object.hasOwn(DEFAULT_LIMITS, key)) throw new RangeError(`Unknown ODS limit ${key}.`);
    if (value === undefined) continue;
    if (!Number.isSafeInteger(value) || value < 1) throw new RangeError(`${key} must be a positive safe integer.`);
    limits[key as keyof OdsLimits] = value;
  }
  return limits;
}

/** Runs a part parser, prefixing XML parse errors with the part name. */
function parsePart<T>(run: () => T, part: string): T {
  try { return run(); }
  catch (error) {
    if (error instanceof PreviewError && error.code === "parse-error" && !error.message.startsWith(`${part}:`)) throw new PreviewError(`${part}: ${error.message}`, "parse-error", { cause: error });
    throw error;
  }
}

const is = (name: ExpandedName, uri: string, local: string) => name.uri === uri && name.local === local;
const attrOf = (element: XmlStartElement, uri: string, local: string) => element.attributes.find(attribute => attribute.uri === uri && attribute.local === local)?.value;
/** A repeat or span count: a positive integer, else 1. */
function count(value: string | undefined): number {
  if (value === undefined) return 1;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= 1 ? parsed : 1;
}

/** Streams content.xml: style blocks and settings are captured, tables are built row by row. */
class ContentReader {
  readonly sheets: ExtendedWorksheet[] = [];
  readonly names: SpreadsheetDefinedName[] = [];
  nullDate = DEFAULT_NULL_DATE;
  precisionAsShown = false;
  formatVersion: string | undefined;
  formulaCells = 0;
  missingResults = 0;
  cellCount = 0;
  #fontFaces: NsXmlNode[] = [];
  #automatic: NsXmlNode | undefined;
  #resolver: OdfStyleResolver | undefined;
  #conditionalStyles = new Set<string>();
  #stack: ExpandedName[] = [];
  #sheet: SheetBuilder | undefined;

  constructor(private readonly stylesRoot: NsXmlNode | undefined, private readonly collector: DiagnosticCollector, private readonly limits: OdsLimits, readonly chartObjects: ReadonlySet<string>) {
    if (stylesRoot) {
      this.#fontFaces.push(...nsChildren(stylesRoot, ODF.office, "font-face-decls"));
      this.#collectConditionalStyles(nsFirst(stylesRoot, ODF.office, "styles"));
    }
  }

  resolver(): OdfStyleResolver {
    this.#resolver ??= createStyleResolver(this.stylesRoot, this.#automatic, this.#fontFaces, this.collector);
    return this.#resolver;
  }

  start(element: XmlStartElement, depth: number): StartAction {
    const name = element.name, parent = this.#stack[this.#stack.length - 1];
    if (depth === 1) {
      if (!is(name, ODF.office, "document-content")) throw new PreviewError("content.xml is not an OpenDocument content stream.", "invalid-file");
      this.formatVersion = attrOf(element, ODF.office, "version");
    } else if (parent && is(parent, ODF.office, "document-content")) {
      if (is(name, ODF.office, "font-face-decls") || is(name, ODF.office, "automatic-styles")) return "capture";
      if (!is(name, ODF.office, "body")) return "skip";
    } else if (parent && is(parent, ODF.office, "body")) {
      if (!is(name, ODF.office, "spreadsheet")) return "skip";
    } else if (parent && is(parent, ODF.office, "spreadsheet")) {
      if (is(name, ODF.table, "calculation-settings") || is(name, ODF.table, "named-expressions")) return "capture";
      if (is(name, ODF.table, "dde-links")) { this.collector.report("odf.external-links"); return "skip"; }
      if (is(name, ODF.table, "data-pilot-tables")) { this.collector.report("odf.unsupported.pivot"); return "skip"; }
      if (!is(name, ODF.table, "table")) return "skip";
      if (this.sheets.length + (this.#sheet ? 1 : 0) >= this.limits.maxSheets) throw new PreviewError("The spreadsheet has more sheets than the configured limit.", "resource-limit");
      this.#sheet = new SheetBuilder(element, this.resolver(), this.collector, this.limits, this);
    } else if (this.#sheet) {
      const action = this.#sheet.start(element, parent);
      if (action !== undefined) return action;
    }
    // streamXml fires `end` itself for self-closing elements, so only open elements are tracked.
    if (!element.selfClosing) this.#stack.push(name);
    return undefined;
  }

  end(name: ExpandedName): void {
    if (this.#stack.length && this.#stack[this.#stack.length - 1] === name) this.#stack.pop();
    if (this.#sheet) {
      if (is(name, ODF.table, "table")) { this.sheets.push(this.#sheet.finish()); this.#sheet = undefined; }
      else this.#sheet.end(name);
    }
  }

  captured(node: NsXmlNode): void {
    if (node.uri === ODF.office && node.local === "font-face-decls") { this.#fontFaces.push(node); return; }
    if (node.uri === ODF.office && node.local === "automatic-styles") { this.#automatic = node; this.#collectConditionalStyles(node); return; }
    if (node.uri === ODF.table && node.local === "calculation-settings") {
      const nullDate = nsAttr(nsFirst(node, ODF.table, "null-date"), ODF.table, "date-value");
      if (nullDate !== undefined) this.nullDate = nullDate;
      this.precisionAsShown = nsAttr(node, ODF.table, "precision-as-shown") === "true";
      return;
    }
    if (node.uri === ODF.table && node.local === "named-expressions") { this.names.push(...readNames(node, this.#sheet?.name)); return; }
    this.#sheet?.captured(node);
  }

  usesConditionalStyle(styleName: string): boolean { return this.#conditionalStyles.has(styleName); }

  isChartObject(href: string | undefined): boolean {
    const path = href?.trim().replace(/^\.\//, "").replace(/\/$/, "");
    return path !== undefined && this.chartObjects.has(path);
  }

  /** Cell styles with style:map conditions on the cell content (conditional formatting, not applied in Phase 1). */
  #collectConditionalStyles(container: NsXmlNode | undefined): void {
    for (const style of container ? nsChildren(container, ODF.style, "style") : []) {
      if (nsAttr(style, ODF.style, "family") !== "table-cell") continue;
      const name = nsAttr(style, ODF.style, "name");
      if (name !== undefined && nsChildren(style, ODF.style, "map").some(map => /^\s*(cell-content|is-true-formula)/.test(nsAttr(map, ODF.style, "condition") ?? ""))) this.#conditionalStyles.add(name);
    }
  }
}

function readNames(node: NsXmlNode, sheet: string | undefined): SpreadsheetDefinedName[] {
  const names: SpreadsheetDefinedName[] = [];
  for (const child of node.children) {
    if (typeof child === "string" || child.uri !== ODF.table) continue;
    const name = nsAttr(child, ODF.table, "name");
    if (name === undefined) continue;
    const base = sheet !== undefined ? { name, sheet } : { name };
    if (child.local === "named-range") { const ref = nsAttr(child, ODF.table, "cell-range-address"); names.push(ref !== undefined ? { ...base, ref } : base); }
    else if (child.local === "named-expression") { const expression = nsAttr(child, ODF.table, "expression"); names.push(expression !== undefined ? { ...base, expression } : base); }
  }
  return names;
}

type PendingRow = { start: number; repeat: number; styleName?: string; visibility?: string; defaultCellStyle?: string; column: number; cells: Array<{ column: number; repeat: number; node: NsXmlNode; covered: boolean }> };
type StyledRun = { top: number; left: number; rows: number; columns: number; style: number };

/** Builds one sheet from table:table events (specification §6.5). */
class SheetBuilder {
  readonly name: string;
  readonly #cells: ExtendedSpreadsheetCell[] = [];
  readonly #columns: ExtendedSpreadsheetColumn[] = [];
  readonly #rowRanges: SpreadsheetRowRange[] = [];
  readonly #columnDefaultStyles: Array<string | undefined> = [];
  readonly #merges: SpreadsheetRange[] = [];
  readonly #hyperlinks: Record<string, string> = {};
  readonly #comments: Record<string, string> = {};
  readonly #styledRuns: StyledRun[] = [];
  readonly #hidden: boolean;
  readonly #tabColor: string | undefined;
  readonly #protected: boolean;
  #nextColumn = 1;
  #nextRow = 1;
  #row: PendingRow | undefined;
  #maxRow = 0;
  #maxColumn = 0;
  #clipped = false;
  #inShapes = 0;

  constructor(element: XmlStartElement, private readonly styles: OdfStyleResolver, private readonly collector: DiagnosticCollector, private readonly limits: OdsLimits, private readonly content: ContentReader) {
    this.name = attrOf(element, ODF.table, "name") ?? `Sheet ${content.sheets.length + 1}`;
    const properties = styles.tableProperties(attrOf(element, ODF.table, "style-name"));
    this.#hidden = properties.hidden;
    this.#tabColor = properties.tabColor;
    this.#protected = attrOf(element, ODF.table, "protected") === "true";
  }

  start(element: XmlStartElement, parent: ExpandedName | undefined): StartAction {
    const name = element.name;
    // Inside table:shapes every element is counted, self-closing ones included: streamXml fires `end` for each.
    if (this.#inShapes) {
      if (is(name, ODF.draw, "frame")) this.collector.report("odf.unsupported.drawings");
      if (is(name, ODF.draw, "object") && this.content.isChartObject(attrOf(element, ODF.xlink, "href"))) this.collector.report("odf.unsupported.charts");
      this.#inShapes++;
      return undefined;
    }
    if (is(name, ODF.table, "shapes")) { this.#inShapes = 1; return undefined; }
    if (is(name, ODF.table, "table-source")) { this.collector.report("odf.external-links"); return "skip"; }
    if (is(name, EXT.calcext, "conditional-formats")) { this.collector.report("odf.unsupported.conditional-formatting"); return "skip"; }
    if (is(name, ODF.table, "named-expressions")) return "capture";
    if (is(name, ODF.table, "table-column")) { this.#column(element); return undefined; }
    if (is(name, ODF.table, "table-row")) { this.#startRow(element); return undefined; }
    if (this.#row && parent && is(parent, ODF.table, "table-row") && (is(name, ODF.table, "table-cell") || is(name, ODF.table, "covered-table-cell"))) return "capture";
    if (is(name, ODF.office, "forms") || is(name, ODF.table, "scenario") || is(name, ODF.table, "title") || is(name, ODF.table, "desc")) return "skip";
    return undefined;
  }

  end(name: ExpandedName): void {
    if (this.#inShapes) { this.#inShapes--; return; }
    if (is(name, ODF.table, "table-row") && this.#row) this.#finishRow();
  }

  captured(node: NsXmlNode): void {
    const row = this.#row;
    if (!row) return;
    const repeat = count(nsAttr(node, ODF.table, "number-columns-repeated"));
    row.cells.push({ column: row.column, repeat, node, covered: node.local === "covered-table-cell" });
    row.column += repeat;
  }

  #column(element: XmlStartElement): void {
    const repeat = count(attrOf(element, ODF.table, "number-columns-repeated"));
    const first = this.#nextColumn, last = first + repeat - 1;
    this.#nextColumn += repeat;
    if (first > MAX_COLUMN) return;
    const visibility = attrOf(element, ODF.table, "visibility"), defaultStyle = attrOf(element, ODF.table, "default-cell-style-name");
    const widthPoints = this.styles.columnWidthPoints(attrOf(element, ODF.table, "style-name"));
    const style = this.styles.cellStyleIndex(defaultStyle);
    const column: ExtendedSpreadsheetColumn = { min: first, max: Math.min(last, MAX_COLUMN) };
    if (widthPoints !== undefined) column.widthPoints = widthPoints;
    if (visibility === "collapse" || visibility === "filter") column.hidden = true;
    if (visibility === "filter") column.filtered = true;
    if (style) column.style = style;
    const previous = this.#columns[this.#columns.length - 1];
    if (previous && previous.max + 1 === column.min && previous.widthPoints === column.widthPoints && previous.hidden === column.hidden && previous.filtered === column.filtered && previous.style === column.style) previous.max = column.max;
    else this.#columns.push(column);
    if (defaultStyle !== undefined) for (let index = first; index <= Math.min(last, MAX_COLUMN); index++) this.#columnDefaultStyles[index] = defaultStyle;
  }

  #startRow(element: XmlStartElement): void {
    const styleName = attrOf(element, ODF.table, "style-name"), visibility = attrOf(element, ODF.table, "visibility"), defaultCellStyle = attrOf(element, ODF.table, "default-cell-style-name");
    this.#row = { start: this.#nextRow, repeat: count(attrOf(element, ODF.table, "number-rows-repeated")), column: 1, cells: [],
      ...(styleName !== undefined ? { styleName } : {}), ...(visibility !== undefined ? { visibility } : {}), ...(defaultCellStyle !== undefined ? { defaultCellStyle } : {}) };
  }

  #finishRow(): void {
    const row = this.#row!;
    this.#row = undefined;
    this.#nextRow += row.repeat;
    if (row.start > MAX_ROW) { if (row.cells.some(cell => hasContent(cell.node))) this.#clip(); return; }
    const lastRow = Math.min(row.start + row.repeat - 1, MAX_ROW);
    if (row.start + row.repeat - 1 > MAX_ROW && row.cells.some(cell => hasContent(cell.node))) this.#clip();
    const rowStyle = this.styles.cellStyleIndex(row.defaultCellStyle);
    this.#addRowRange({ min: row.start, max: lastRow, ...this.#rowHeight(row.styleName), ...(row.visibility === "collapse" || row.visibility === "filter" ? { hidden: true } : {}), ...(row.visibility === "filter" ? { filtered: true } : {}), ...(rowStyle ? { style: rowStyle } : {}) });
    for (const cell of row.cells) this.#cellRun(cell, row, lastRow);
  }

  #rowHeight(styleName: string | undefined): { height?: number } {
    const height = this.styles.rowProperties(styleName).heightPt;
    return height !== undefined ? { height } : {};
  }

  #addRowRange(range: SpreadsheetRowRange): void {
    if (range.height === undefined && !range.hidden && range.style === undefined) return;
    const previous = this.#rowRanges[this.#rowRanges.length - 1];
    if (previous && previous.max + 1 === range.min && previous.height === range.height && previous.hidden === range.hidden && previous.filtered === range.filtered && previous.style === range.style) previous.max = range.max;
    else this.#rowRanges.push(range);
  }

  /** Materialises one table:table-cell (repeated across columns and the row's repeated rows). */
  #cellRun(cell: PendingRow["cells"][number], row: PendingRow, lastRow: number): void {
    const { node, covered } = cell;
    const firstColumn = cell.column, lastColumn = Math.min(cell.column + cell.repeat - 1, MAX_COLUMN);
    const content = hasContent(node), ownStyle = nsAttr(node, ODF.table, "style-name");
    if (firstColumn > MAX_COLUMN || cell.column + cell.repeat - 1 > MAX_COLUMN) { if (content) this.#clip(); if (firstColumn > MAX_COLUMN) return; }
    this.#countDrawings(node);
    const inheritedStyle = (column: number) => row.defaultCellStyle ?? this.#columnDefaultStyles[column];
    if (ownStyle !== undefined && this.content.usesConditionalStyle(ownStyle)) this.collector.report("odf.unsupported.conditional-formatting");
    const comment = annotationText(node);
    if (!content) {
      const style = this.styles.cellStyleIndex(ownStyle);
      if (comment !== undefined) for (let r = row.start; r <= lastRow; r++) for (let c = firstColumn; c <= lastColumn; c++) this.#comments[`${columnName(c)}${r}`] = comment;
      if (ownStyle !== undefined && style !== 0 && style !== this.styles.cellStyleIndex(inheritedStyle(firstColumn))) {
        this.#styledRuns.push({ top: row.start, left: firstColumn, rows: lastRow - row.start + 1, columns: lastColumn - firstColumn + 1, style });
      }
      return;
    }
    const copies = (lastRow - row.start + 1) * (lastColumn - firstColumn + 1);
    this.content.cellCount += copies;
    if (this.content.cellCount > this.limits.maxCells) throw new PreviewError("The spreadsheet exceeds the configured cell limit.", "resource-limit");
    const saved = savedCellText(node);
    if (saved !== undefined && saved.length > this.limits.maxCellTextLength) throw new PreviewError("A cell's text exceeds the configured length limit.", "resource-limit");
    const columnSpan = covered ? 1 : count(nsAttr(node, ODF.table, "number-columns-spanned")), rowSpan = covered ? 1 : count(nsAttr(node, ODF.table, "number-rows-spanned"));
    const link = cellHyperlink(node), runs = cellRuns(node, this.styles);
    const value = readCellValue(node, saved, this.collector, this.content.nullDate);
    const formula = readFormula(node, value, saved, this.collector, row.start, firstColumn);
    if (formula) { this.content.formulaCells += copies; if (formula.result === "missing") this.content.missingResults += copies; }
    const display = displayFor(value, saved);
    if (display.source === "raw" && isNumericKind(value)) this.collector.report("odf.display.raw", { count: copies });
    const data: SpreadsheetCellData = { value, display, ...(formula ? { formula } : {}), ...(covered ? { covered: true } : {}) };
    if (covered) this.collector.report("odf.covered-cell-content", { count: copies });
    for (let r = row.start; r <= lastRow; r++) {
      for (let c = firstColumn; c <= lastColumn; c++) {
        const ref = `${columnName(c)}${r}`, style = this.styles.cellStyleIndex(ownStyle ?? inheritedStyle(c));
        this.#cells.push(legacyCell(ref, data, style, runs));
        if (link) this.#hyperlinks[ref] = link;
        if (comment !== undefined) this.#comments[ref] = comment;
        if (!covered && (columnSpan > 1 || rowSpan > 1)) {
          const merge = { top: r, left: c, bottom: Math.min(r + rowSpan - 1, MAX_ROW), right: Math.min(c + columnSpan - 1, MAX_COLUMN) };
          this.#merges.push(merge);
          this.#maxRow = Math.max(this.#maxRow, merge.bottom); this.#maxColumn = Math.max(this.#maxColumn, merge.right);
        }
      }
    }
    this.#maxRow = Math.max(this.#maxRow, lastRow); this.#maxColumn = Math.max(this.#maxColumn, lastColumn);
  }

  /** Drawings anchored in a cell (draw:frame inside table:table-cell). */
  #countDrawings(node: NsXmlNode): void {
    const walk = (element: NsXmlNode) => {
      for (const child of element.children) {
        if (typeof child === "string") continue;
        if (child.uri === ODF.draw && child.local === "frame") this.collector.report("odf.unsupported.drawings");
        if (child.uri === ODF.draw && child.local === "object" && this.content.isChartObject(nsAttr(child, ODF.xlink, "href"))) this.collector.report("odf.unsupported.charts");
        walk(child);
      }
    };
    walk(node);
  }

  #clip(): void {
    if (!this.#clipped) this.collector.report("odf.bounds.clipped");
    this.#clipped = true;
  }

  finish(): ExtendedWorksheet {
    if (this.#row) this.#finishRow();
    // Styled empty cells inside the drawn grid (specification §6.5 item 3).
    const rowWindow = Math.max(this.#maxRow + RENDER_MARGIN_ROWS, RENDER_MIN_ROWS), columnWindow = Math.max(this.#maxColumn + RENDER_MARGIN_COLUMNS, RENDER_MIN_COLUMNS);
    for (const run of this.#styledRuns) {
      const bottom = Math.min(run.top + run.rows - 1, rowWindow), right = Math.min(run.left + run.columns - 1, columnWindow);
      if (run.top > bottom || run.left > right) continue;
      const copies = (bottom - run.top + 1) * (right - run.left + 1);
      this.content.cellCount += copies;
      if (this.content.cellCount > this.limits.maxCells) throw new PreviewError("The spreadsheet exceeds the configured cell limit.", "resource-limit");
      for (let r = run.top; r <= bottom; r++) for (let c = run.left; c <= right; c++) this.#cells.push({ ref: `${columnName(c)}${r}`, value: "", style: run.style });
    }
    this.#reportHidden();
    const sheet: ExtendedWorksheet = { name: this.name, cells: this.#cells, maxRow: this.#maxRow, maxColumn: this.#maxColumn, defaultColumnWidthPoints: DEFAULT_COLUMN_WIDTH_POINTS, defaultRowHeight: DEFAULT_ROW_HEIGHT_POINTS };
    if (this.#hidden) sheet.hidden = true;
    if (this.#tabColor) sheet.tabColor = this.#tabColor;
    if (this.#protected) sheet.protected = true;
    if (this.#columns.length) sheet.columns = this.#columns;
    if (this.#rowRanges.length) sheet.rowRanges = this.#rowRanges;
    if (this.#merges.length) sheet.merges = this.#merges;
    if (Object.keys(this.#hyperlinks).length) sheet.hyperlinks = this.#hyperlinks;
    if (Object.keys(this.#comments).length) sheet.comments = this.#comments;
    return sheet;
  }

  /** Hidden and filtered rows and hidden columns within the used range (specification §6.5 item 6). */
  #reportHidden(): void {
    let hiddenRows = 0, filteredRows = 0, hiddenColumns = 0;
    for (const range of this.#rowRanges) {
      if (!range.hidden) continue;
      const rows = Math.max(0, Math.min(range.max, this.#maxRow) - range.min + 1);
      if (range.filtered) filteredRows += rows; else hiddenRows += rows;
    }
    for (const column of this.#columns) if (column.hidden) hiddenColumns += Math.max(0, Math.min(column.max, this.#maxColumn) - column.min + 1);
    if (hiddenRows) this.collector.report("odf.hidden.rows", { count: hiddenRows });
    if (filteredRows) this.collector.report("odf.hidden.filtered-rows", { count: filteredRows });
    if (hiddenColumns) this.collector.report("odf.hidden.columns", { count: hiddenColumns });
  }
}

/** A cell has content when it has a value type, saved text or a formula. Styles and comments alone are not content. */
function hasContent(node: NsXmlNode): boolean {
  if (nsAttr(node, ODF.office, "value-type") !== undefined || nsAttr(node, ODF.table, "formula") !== undefined) return true;
  return node.children.some(child => typeof child !== "string" && child.uri === ODF.text && (child.local === "p" || child.local === "h"));
}

const isNumericKind = (value: SpreadsheetCellValue) => value.kind === "number" || value.kind === "percentage" || value.kind === "currency" || value.kind === "date" || value.kind === "time";

/** Display selection for M11 (specification §6.3): saved text first, never formatted with the viewer's locale. */
function displayFor(value: SpreadsheetCellValue, saved: string | undefined): SpreadsheetDisplay {
  if (saved !== undefined) return { text: saved, source: "file" };
  switch (value.kind) {
    case "empty": case "invalid": return { text: "", source: "raw" };
    case "string": return { text: value.text, source: "file" };
    case "boolean": return { text: value.boolean ? "TRUE" : "FALSE", source: "raw" };
    case "error": return { text: value.code, source: "raw" };
    default: return { text: value.lexical, source: "raw" };
  }
}

/** The cell's formula; `top`/`left` place a matrix (array) formula's range on the sheet. */
function readFormula(node: NsXmlNode, value: SpreadsheetCellValue, saved: string | undefined, collector: DiagnosticCollector, top: number, left: number): SpreadsheetFormula | undefined {
  const text = nsAttr(node, ODF.table, "formula");
  if (text === undefined) return undefined;
  const prefix = /^[^:=]+:=/.test(text) ? nsValuePrefix(node, text) : undefined;
  const dialect = prefix ? FORMULA_NAMESPACES.get(prefix.uri) : undefined;
  if (!dialect) collector.report("odf.formula.unknown-dialect");
  const columns = nsAttr(node, ODF.table, "number-matrix-columns-spanned"), rows = nsAttr(node, ODF.table, "number-matrix-rows-spanned");
  return {
    text, expression: prefix && dialect ? prefix.local : text, dialect: dialect ?? "unknown",
    ...(prefix ? { namespace: prefix.uri } : {}),
    result: value.kind !== "empty" || saved !== undefined ? "cached" : "missing",
    ...(columns !== undefined || rows !== undefined ? { matrix: { top, left, bottom: top + count(rows) - 1, right: left + count(columns) - 1 } } : {}),
  };
}

/** Legacy SpreadsheetCell fields (specification §6.4), so the existing grid renderer draws ODS cells unchanged. */
function legacyCell(ref: string, data: SpreadsheetCellData, style: number, runs: ReturnType<typeof cellRuns>): ExtendedSpreadsheetCell {
  const cell: ExtendedSpreadsheetCell = { ref, value: "", data };
  if (style) cell.style = style;
  if (data.formula) cell.formula = data.formula.expression;
  if (data.formula?.result === "missing") return cell;
  const value = data.value;
  switch (value.kind) {
    case "empty": return cell;
    case "number": case "percentage": case "currency": cell.value = String(value.number); cell.type = "n"; break;
    case "date": if (value.serial !== undefined) { cell.value = String(value.serial); cell.type = "n"; } else { cell.value = value.lexical; cell.type = "str"; } break;
    case "time": if (value.days !== undefined) { cell.value = String(value.days); cell.type = "n"; } else { cell.value = value.lexical; cell.type = "str"; } break;
    case "boolean": cell.value = value.boolean ? "1" : "0"; cell.type = "b"; break;
    case "string": cell.value = value.text; cell.type = "str"; break;
    case "error": cell.value = value.code; cell.type = "e"; cell.display = value.code; return cell;
    case "invalid": cell.value = data.display.source === "file" ? data.display.text : ""; cell.type = "str"; break;
  }
  if (value.kind !== "invalid" || data.display.source === "file") cell.display = data.display.text;
  if (runs && cell.value === cell.display) cell.runs = runs;
  return cell;
}
