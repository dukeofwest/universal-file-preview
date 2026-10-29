import { PreviewError } from "../errors.js";
import { ZipArchive } from "../archive/zip-reader.js";
import { parseXml, elements, firstElement, textContent, type XmlNode } from "../xml/xml-parser.js";
import type { SpreadsheetAnchor, SpreadsheetBorderEdge, SpreadsheetBorders, SpreadsheetCell, SpreadsheetConditional, SpreadsheetDrawing, SpreadsheetFill, SpreadsheetFont, SpreadsheetRange, SpreadsheetRow, SpreadsheetRun, SpreadsheetSlicer, SpreadsheetSlicerStyle, SpreadsheetSparkline, SpreadsheetStyle, SpreadsheetTable, Workbook, Worksheet } from "../model.js";
import { builtinNumberFormat, dateToSerial, formatValue, indexedColors as numberFormatColors, type FormattedValue } from "./xlsx-number-format.js";
import { evaluateFormula, FormulaError, isTruthy, type FormulaContext, type FormulaScalar } from "./xlsx-formula.js";
import { builtinTableStyle, type TableStyleColor } from "./xlsx-table-styles.js";
import { parseDrawingShapes } from "./ooxml.js";
import { tiffToPng } from "./tiff.js";

export type XlsxParseOptions = { locale?: string; today?: number };

const decodeXml = (archive: ZipArchive, path: string): XmlNode => parseXml(new TextDecoder("utf-8", { fatal: true }).decode(archive.read(path)));
const child = (node: XmlNode | undefined, name: string) => node ? firstElement(node, name) : undefined;
const attr = (node: XmlNode | undefined, name: string): string | undefined => node ? node.attributes[name] ?? Object.entries(node.attributes).find(([key]) => key.endsWith(`:${name}`))?.[1] : undefined;
const flag = (node: XmlNode | undefined, name: string, fallback = false): boolean => { const value = attr(node, name); return value === undefined ? fallback : value === "1" || value === "true"; };
/** OOXML boolean elements such as <b/> or <b val="0"/>. */
const on = (node: XmlNode | undefined): boolean | undefined => node ? !/^(0|false|none)$/i.test(attr(node, "val") ?? "1") : undefined;
const MAX_CF_CELLS = 250_000;

type Relationship = { target: string; type: string; external: boolean };
type Theme = { colors: string[]; majorFont?: string; minorFont?: string };
type StyleSheet = { cellStyles: SpreadsheetStyle[]; dxfs: SpreadsheetStyle[]; tableStyles: Map<string, Record<string, SpreadsheetStyle & { size?: number }>>; slicerStyles: Map<string, Record<string, SpreadsheetStyle>>; defaultSlicerStyle?: string; defaultFont: SpreadsheetFont; fontIds: number[] };
type CellValue = FormulaScalar;
/** Table data needed by slicers, which may sit on another sheet than their table. */
type TableInfo = { sheet: Worksheet; range: SpreadsheetRange; headerRows: number; totalsRows: number; columnIds: number[]; filters: Map<number, { values: Set<string>; blank: boolean }> };
type PendingSlicer = { drawing: SpreadsheetSlicer; cache: string };

export function parseXlsx(bytes: Uint8Array, name: string, options: XlsxParseOptions = {}): Workbook {
  const locale = options.locale ?? "en-US";
  const zip = ZipArchive.open(bytes), workbook = decodeXml(zip, "xl/workbook.xml");
  const relationships = readRelationships(zip, "xl/workbook.xml");
  const themePath = [...relationships.values()].find(rel => rel.type.endsWith("/theme"))?.target ?? "xl/theme/theme1.xml";
  const themeRoot = zip.has(themePath) ? decodeXml(zip, themePath) : undefined, theme = readTheme(themeRoot);
  const stylesPath = [...relationships.values()].find(rel => rel.type.endsWith("/styles"))?.target ?? "xl/styles.xml";
  const stylesRoot = zip.has(stylesPath) ? decodeXml(zip, stylesPath) : undefined;
  const palette = readIndexedPalette(stylesRoot);
  const color = (node: XmlNode | undefined, fallback?: string) => readColor(node, theme, palette) ?? fallback;
  const styles = readStyles(stylesRoot, theme, palette, locale);
  const sharedPath = [...relationships.values()].find(rel => rel.type.endsWith("/sharedStrings"))?.target ?? "xl/sharedStrings.xml";
  const shared = zip.has(sharedPath) ? elements(decodeXml(zip, sharedPath), "si").map(si => readRichText(si, theme, palette)) : [];
  const date1904 = flag(child(workbook, "workbookPr"), "date1904");
  const formatFor = (styleIndex: number | undefined) => styles.cellStyles[styleIndex ?? 0]?.numberFormat ?? "General";

  const definedNames = elements(child(workbook, "definedNames") ?? workbook, "definedName").map(node => ({ name: attr(node, "name") ?? "", scope: attr(node, "localSheetId"), formula: textContent(node) }));
  const sheetNodes = elements(child(workbook, "sheets") ?? workbook, "sheet");
  type Parsed = { sheet: Worksheet; values: Map<number, CellValue>; formulas: Map<number, { formula: string; origin: { row: number; column: number } }>; root: XmlNode; path: string; index: number };
  const parsed: Parsed[] = [];
  sheetNodes.forEach((sheetNode, index) => {
    const sheetName = attr(sheetNode, "name") ?? `Sheet ${index + 1}`;
    const relId = Object.entries(sheetNode.attributes).find(([key]) => key.endsWith(":id"))?.[1];
    const rel = relationships.get(relId ?? "");
    if (!rel || rel.external || !zip.has(rel.target)) return;
    const root = decodeXml(zip, rel.target);
    if (root.localName !== "worksheet") return;
    const state = attr(sheetNode, "state");
    const { sheet, values, formulas } = readWorksheet(root, sheetName, shared, color, formatFor, locale, date1904);
    if (state === "hidden" || state === "veryHidden") sheet.hidden = true;
    parsed.push({ sheet, values, formulas, root, path: rel.target, index });
  });
  if (!parsed.length) throw new PreviewError("XLSX workbook contains no readable worksheets.", "invalid-file");

  // Formula context spanning all sheets, so conditional formats and uncached formulas can reference any sheet.
  const bySheet = new Map(parsed.map(entry => [entry.sheet.name.toLowerCase(), entry]));
  const evaluating = new Set<string>();
  const cellValue = (sheetName: string, row: number, column: number): CellValue => {
    const entry = bySheet.get(sheetName.toLowerCase());
    if (!entry) return new FormulaError("#REF!");
    const key = row * 16_384 + column;
    const value = entry.values.get(key);
    if (value !== undefined) return value;
    const pending = entry.formulas.get(key);
    if (!pending) return null;
    const guard = `${sheetName}!${key}`;
    if (evaluating.has(guard)) return new FormulaError("#REF!");
    evaluating.add(guard);
    const result = evaluateFormula(pending.formula, contextFor(sheetName, row, column), pending.origin);
    evaluating.delete(guard);
    const scalar = typeof result === "object" && result !== null && !(result instanceof FormulaError) ? new FormulaError("#VALUE!") : result as CellValue;
    entry.values.set(key, scalar ?? 0);
    return scalar ?? 0;
  };
  const nameFor = (nameText: string, sheetName: string): string | undefined => {
    const sheetIndex = parsed.find(entry => entry.sheet.name.toLowerCase() === sheetName.toLowerCase())?.index;
    const lower = nameText.toLowerCase();
    return (definedNames.find(entry => entry.name.toLowerCase() === lower && entry.scope !== undefined && Number(entry.scope) === sheetIndex)
      ?? definedNames.find(entry => entry.name.toLowerCase() === lower && entry.scope === undefined))?.formula;
  };
  const contextFor = (sheetName: string, row: number, column: number): FormulaContext => ({ sheet: sheetName, row, column, cell: cellValue, name: nameFor, ...(options.today !== undefined ? { today: options.today } : {}) });

  for (const entry of parsed) {
    // Formula cells saved without cached values (common for generated files) are computed where supported.
    const cellsByRef = entry.formulas.size ? new Map(entry.sheet.cells.map(cell => [cell.ref.toUpperCase(), cell])) : undefined;
    for (const key of entry.formulas.keys()) {
      const row = Math.floor(key / 16_384), column = key % 16_384;
      const value = cellValue(entry.sheet.name, row, column);
      const cell = cellsByRef?.get(`${columnName(column)}${row}`);
      if (!cell || value === null || value instanceof FormulaError && value.code === "#NAME?") continue;
      const formatted = displayFor(value, formatFor(cell.style), locale, date1904);
      cell.value = value instanceof FormulaError ? value.code : typeof value === "boolean" ? (value ? "1" : "0") : String(value);
      if (typeof value === "string") cell.type = "str"; else if (typeof value === "boolean") cell.type = "b"; else if (value instanceof FormulaError) cell.type = "e";
      applyDisplay(cell, formatted);
    }
  }
  const tablesById = new Map<number, TableInfo>(), pendingSlicers: PendingSlicer[] = [];
  for (const entry of parsed) {
    const rels = readRelationships(zip, entry.path);
    readSheetExtras(zip, entry.root, entry.sheet, rels, styles, theme, themeRoot, tablesById, pendingSlicers);
    const conditional = evaluateConditionalFormats(entry.root, entry.sheet, styles, theme, palette, (row, column) => cellValue(entry.sheet.name, row, column), (row, column) => contextFor(entry.sheet.name, row, column), locale, date1904);
    if (Object.keys(conditional).length) entry.sheet.conditional = conditional;
    const sparklines = readSparklines(entry.root, entry.sheet.name, cellValue, theme, palette);
    if (sparklines.length) entry.sheet.sparklines = sparklines;
  }
  if (pendingSlicers.length) resolveSlicers(zip, workbook, relationships, pendingSlicers, tablesById, locale);

  const view = child(child(workbook, "bookViews"), "workbookView");
  const activeTab = Number(attr(view, "activeTab") ?? 0);
  const activeSheet = Math.max(0, parsed.findIndex(entry => entry.index === activeTab));
  return { kind: "workbook", sheets: parsed.map(entry => entry.sheet), title: name, styles: styles.cellStyles, defaultFont: styles.defaultFont, ...(activeSheet ? { activeSheet } : {}), ...(date1904 ? { date1904 } : {}) };
}

function applyDisplay(cell: SpreadsheetCell, formatted: FormattedValue): void {
  cell.display = formatted.text;
  if (formatted.color) cell.displayColor = formatted.color; else delete cell.displayColor;
  if (formatted.repeat) cell.repeat = formatted.repeat; else delete cell.repeat;
}

function displayFor(value: CellValue, format: string, locale: string, date1904: boolean): FormattedValue {
  if (value === null) return { text: "" };
  if (value instanceof FormulaError) return { text: value.code };
  return formatValue(value, format, { locale, date1904 });
}

function readWorksheet(root: XmlNode, name: string, shared: Array<{ text: string; runs?: SpreadsheetRun[] }>, color: (node: XmlNode | undefined, fallback?: string) => string | undefined, formatFor: (style: number | undefined) => string, locale: string, date1904: boolean) {
  const cells: SpreadsheetCell[] = [], values = new Map<number, CellValue>(), formulas = new Map<number, { formula: string; origin: { row: number; column: number } }>();
  const sharedFormulas = new Map<string, { formula: string; origin: { row: number; column: number } }>();
  const rows: Record<number, SpreadsheetRow> = {};
  let maxRow = 0, maxColumn = 0, rowNumber = 0;
  for (const row of elements(child(root, "sheetData") ?? root, "row")) {
    rowNumber = Number(attr(row, "r") ?? rowNumber + 1);
    if (!Number.isSafeInteger(rowNumber) || rowNumber < 1 || rowNumber > 1_048_576) continue;
    const height = Number(attr(row, "ht")), info: SpreadsheetRow = {};
    if (attr(row, "ht") !== undefined && Number.isFinite(height)) info.height = height;
    if (flag(row, "hidden")) info.hidden = true;
    if (flag(row, "customFormat") && attr(row, "s") !== undefined) info.style = Number(attr(row, "s"));
    if (Object.keys(info).length) rows[rowNumber] = info;
    let columnNumber = 0;
    for (const cell of elements(row, "c")) {
      const ref = attr(cell, "r");
      const position = ref ? parseCellReference(ref) : { row: rowNumber, column: columnNumber + 1 };
      if (!position.row || !position.column) continue;
      columnNumber = position.column;
      const type = attr(cell, "t"), styleIndex = attr(cell, "s") !== undefined ? Number(attr(cell, "s")) : undefined;
      const formulaNode = child(cell, "f"), valueNode = child(cell, "v");
      const output: SpreadsheetCell = { ref: ref ?? `${columnName(position.column)}${position.row}`, value: valueNode ? textContent(valueNode) : "" };
      let runs: SpreadsheetRun[] | undefined;
      if (type === "s") { const entry = shared[Number(output.value)]; output.value = entry?.text ?? ""; runs = entry?.runs; }
      else if (type === "inlineStr") { const entry = child(cell, "is") ? readRichText(child(cell, "is")!, undefined, undefined, color) : { text: "" }; output.value = entry.text; runs = entry.runs; }
      if (formulaNode) {
        output.formula = textContent(formulaNode);
        const si = attr(formulaNode, "si");
        if (attr(formulaNode, "t") === "shared" && si !== undefined) {
          if (output.formula) sharedFormulas.set(si, { formula: output.formula, origin: position });
        }
        if (!valueNode) {
          const source = output.formula ? { formula: output.formula, origin: position } : si !== undefined ? sharedFormulas.get(si) : undefined;
          if (source) formulas.set(position.row * 16_384 + position.column, source);
        }
      }
      if (type) output.type = type;
      if (styleIndex !== undefined && Number.isFinite(styleIndex)) output.style = styleIndex;
      if (runs && runs.length) output.runs = runs;
      const key = position.row * 16_384 + position.column;
      if (valueNode || type === "inlineStr") {
        let value: CellValue;
        if (type === "s" || type === "inlineStr" || type === "str") value = output.value;
        else if (type === "b") value = output.value === "1" || output.value.toLowerCase() === "true";
        else if (type === "e") value = new FormulaError(output.value);
        else if (type === "d") { const date = new Date(output.value); value = Number.isNaN(date.valueOf()) ? output.value : dateToSerial(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date1904) + (date.getUTCHours() * 3600 + date.getUTCMinutes() * 60 + date.getUTCSeconds()) / 86400; }
        else { const number = Number(output.value); value = output.value.trim() !== "" && Number.isFinite(number) ? number : output.value; }
        values.set(key, value);
        if (!(value === "" && !formulaNode)) applyDisplay(output, displayFor(value, formatFor(styleIndex), locale, date1904));
      }
      cells.push(output);
      maxRow = Math.max(maxRow, position.row); maxColumn = Math.max(maxColumn, position.column);
    }
  }
  const sheet: Worksheet = { name, cells, maxRow, maxColumn };
  if (Object.keys(rows).length) sheet.rows = rows;
  // Column widths/styles, default sizes, and view settings.
  const columns = elements(child(root, "cols") ?? root, "col").map(col => {
    const min = Number(attr(col, "min")), max = Number(attr(col, "max")), width = Number(attr(col, "width"));
    return { min, max: Math.min(max, 16_384), ...(attr(col, "width") !== undefined && Number.isFinite(width) ? { width } : {}), ...(flag(col, "hidden") ? { hidden: true } : {}), ...(attr(col, "style") !== undefined ? { style: Number(attr(col, "style")) } : {}) };
  }).filter(col => Number.isFinite(col.min) && Number.isFinite(col.max) && col.min >= 1 && col.max >= col.min);
  if (columns.length) sheet.columns = columns;
  const format = child(root, "sheetFormatPr");
  const defaultRowHeight = Number(attr(format, "defaultRowHeight")), defaultColumnWidth = Number(attr(format, "defaultColWidth")), baseColumnWidth = Number(attr(format, "baseColWidth"));
  if (Number.isFinite(defaultRowHeight) && defaultRowHeight > 0) sheet.defaultRowHeight = defaultRowHeight;
  if (Number.isFinite(defaultColumnWidth) && defaultColumnWidth > 0) sheet.defaultColumnWidth = defaultColumnWidth;
  else if (Number.isFinite(baseColumnWidth) && baseColumnWidth > 0) sheet.defaultColumnWidth = -baseColumnWidth; // negative marks a base width (padding added by the renderer)
  if (flag(format, "zeroHeight")) sheet.defaultRowHeight = 0;
  const view = child(child(root, "sheetViews"), "sheetView");
  if (view) {
    if (attr(view, "showGridLines") !== undefined) sheet.showGridLines = flag(view, "showGridLines", true);
    if (attr(view, "showRowColHeaders") !== undefined) sheet.showHeaders = flag(view, "showRowColHeaders", true);
    const zoom = Number(attr(view, "zoomScale"));
    if (Number.isFinite(zoom) && zoom >= 10 && zoom <= 400 && zoom !== 100) sheet.zoom = zoom / 100;
    const topLeft = attr(view, "topLeftCell");
    const pane = child(view, "pane");
    if (pane && /^frozen/.test(attr(pane, "state") ?? "")) {
      const rowsFrozen = Math.max(0, Math.floor(Number(attr(pane, "ySplit") ?? 0))), columnsFrozen = Math.max(0, Math.floor(Number(attr(pane, "xSplit") ?? 0)));
      if (rowsFrozen || columnsFrozen) sheet.frozen = { rows: rowsFrozen, columns: columnsFrozen };
    } else if (topLeft && topLeft !== "A1" && parseCellReference(topLeft).row) sheet.topLeftCell = topLeft;
  }
  const tab = child(child(root, "sheetPr"), "tabColor");
  const tabColor = color(tab);
  if (tabColor) sheet.tabColor = tabColor;
  const merges = elements(child(root, "mergeCells") ?? root, "mergeCell").map(node => parseRange(attr(node, "ref") ?? "")).filter((range): range is SpreadsheetRange => !!range && (range.bottom > range.top || range.right > range.left));
  if (merges.length) sheet.merges = merges;
  return { sheet, values, formulas };
}

function readSheetExtras(zip: ZipArchive, root: XmlNode, sheet: Worksheet, rels: Map<string, Relationship>, styles: StyleSheet, theme: Theme, themeRoot: XmlNode | undefined, tablesById: Map<number, TableInfo>, pendingSlicers: PendingSlicer[]): void {
  const relFor = (node: XmlNode | undefined) => rels.get(Object.entries(node?.attributes ?? {}).find(([key]) => key.endsWith(":id"))?.[1] ?? "");
  // Hyperlinks: only web and mail targets become clickable.
  const hyperlinks: Record<string, string> = {};
  for (const link of elements(child(root, "hyperlinks") ?? root, "hyperlink")) {
    const ref = attr(link, "ref"), rel = relFor(link), target = rel?.target;
    if (!ref || !target || !/^(https?:|mailto:)/i.test(target)) continue;
    const range = parseRange(ref);
    if (range) for (let row = range.top; row <= Math.min(range.bottom, range.top + 1000); row++) for (let column = range.left; column <= Math.min(range.right, range.left + 100); column++) hyperlinks[`${columnName(column)}${row}`] = target;
  }
  if (Object.keys(hyperlinks).length) sheet.hyperlinks = hyperlinks;
  // Notes: Excel marks commented cells with a red corner triangle.
  const comments: Record<string, string> = {};
  for (const rel of rels.values()) {
    if (rel.external || !rel.type.endsWith("/comments") || !zip.has(rel.target)) continue;
    for (const comment of elements(child(decodeXml(zip, rel.target), "commentList") ?? { name: "", localName: "", attributes: {}, children: [] }, "comment")) {
      const ref = attr(comment, "ref");
      if (ref) comments[ref] = readRichText(child(comment, "text") ?? comment).text;
    }
  }
  if (Object.keys(comments).length) sheet.comments = comments;
  // Data validation input messages (Excel shows them next to the selected cell).
  const inputMessages = elements(child(root, "dataValidations") ?? root, "dataValidation").flatMap(node => {
    const text = attr(node, "prompt"), title = attr(node, "promptTitle");
    const ranges = (attr(node, "sqref") ?? "").trim().split(/\s+/).map(parseRange).filter((range): range is SpreadsheetRange => !!range);
    return flag(node, "showInputMessage") && (text || title) && ranges.length ? [{ ranges, ...(title ? { title } : {}), text: text ?? "" }] : [];
  });
  if (inputMessages.length) sheet.inputMessages = inputMessages;
  const filterButtons = new Set<string>();
  // Filter columns can hide their drop-down button (hiddenButton, or showButton="0").
  const addFilterButtons = (filter: XmlNode | undefined, range: SpreadsheetRange) => {
    const hidden = new Set((filter ? elements(filter, "filterColumn") : []).filter(column => flag(column, "hiddenButton") || !flag(column, "showButton", true)).map(column => Number(attr(column, "colId"))));
    for (let column = range.left; column <= range.right; column++) if (!hidden.has(column - range.left)) filterButtons.add(`${columnName(column)}${range.top}`);
  };
  const sheetFilter = parseRange(attr(child(root, "autoFilter"), "ref") ?? "");
  if (sheetFilter) addFilterButtons(child(root, "autoFilter"), sheetFilter);
  // Tables ("Format as Table") with built-in or custom table styles.
  const tables: SpreadsheetTable[] = [];
  for (const part of elements(child(root, "tableParts") ?? root, "tablePart")) {
    const rel = relFor(part);
    if (!rel || rel.external || !zip.has(rel.target)) continue;
    const table = decodeXml(zip, rel.target), range = parseRange(attr(table, "ref") ?? "");
    if (!range) continue;
    const headerRows = Number(attr(table, "headerRowCount") ?? 1), totalsRows = Number(attr(table, "totalsRowCount") ?? 0);
    const info = child(table, "tableStyleInfo"), styleName = attr(info, "name");
    const elementsByType: SpreadsheetTable["elements"] = {};
    const custom = styleName ? styles.tableStyles.get(styleName) : undefined;
    const builtin = styleName && !custom ? builtinTableStyle(styleName) : undefined;
    if (custom) Object.assign(elementsByType, custom);
    else if (builtin) for (const [key, element] of Object.entries(builtin)) {
      const style: SpreadsheetStyle & { size?: number } = {};
      const resolve = (value: TableStyleColor) => "rgb" in value ? value.rgb : applyTint(theme.colors[value.theme] ?? "#000000", value.tint);
      if (element.fill) style.fill = { type: "pattern", pattern: "solid", foreground: resolve(element.fill) };
      if (element.color || element.bold || element.italic) style.font = { ...(element.color ? { color: resolve(element.color) } : {}), ...(element.bold ? { bold: true } : {}), ...(element.italic ? { italic: true } : {}) };
      if (element.borders) {
        const borders: SpreadsheetBorders & Partial<Record<"insideV" | "insideH", SpreadsheetBorderEdge>> = {};
        for (const [side, border] of Object.entries(element.borders)) (borders as Record<string, SpreadsheetBorderEdge>)[side] = { style: border.style, color: resolve(border.color) };
        style.borders = borders;
      }
      if (element.size) style.size = element.size;
      elementsByType[key] = style;
    }
    if (headerRows > 0 && child(table, "autoFilter")) addFilterButtons(child(table, "autoFilter"), range);
    const tableId = Number(attr(table, "id"));
    if (Number.isFinite(tableId)) {
      const filters: TableInfo["filters"] = new Map();
      for (const column of elements(child(table, "autoFilter") ?? table, "filterColumn")) {
        const list = child(column, "filters");
        if (list) filters.set(Number(attr(column, "colId")), { values: new Set(elements(list, "filter").map(filter => (attr(filter, "val") ?? "").toLowerCase())), blank: flag(list, "blank") });
      }
      const columnIds = elements(child(table, "tableColumns") ?? table, "tableColumn").map(column => Number(attr(column, "id")));
      tablesById.set(tableId, { sheet, range, headerRows, totalsRows, columnIds, filters });
    }
    tables.push({ range, headerRows, totalsRows, elements: elementsByType, showRowStripes: flag(info, "showRowStripes"), showColumnStripes: flag(info, "showColumnStripes"), showFirstColumn: flag(info, "showFirstColumn"), showLastColumn: flag(info, "showLastColumn") });
  }
  if (tables.length) sheet.tables = tables;
  if (filterButtons.size) sheet.filterButtons = [...filterButtons];
  // Drawings: pictures, shapes and groups anchored to cells.
  const drawingRel = relFor(child(root, "drawing"));
  if (drawingRel && !drawingRel.external && zip.has(drawingRel.target)) {
    const drawings = readDrawings(zip, drawingRel.target, themeRoot);
    // Slicer frames name a slicer of this sheet; its part gives caption, layout, style and cache (items are resolved once all tables are read).
    const definitions = new Map<string, XmlNode>();
    for (const rel of rels.values()) if (!rel.external && rel.type.endsWith("/slicer") && zip.has(rel.target)) for (const node of elements(decodeXml(zip, rel.target), "slicer")) definitions.set(attr(node, "name") ?? "", node);
    for (const drawing of drawings) {
      if (drawing.kind !== "slicer") continue;
      const definition = definitions.get(drawing.name), cache = attr(definition, "cache");
      if (!definition || !cache) continue;
      const caption = attr(definition, "caption") ?? drawing.name, columns = Number(attr(definition, "columnCount") ?? 1), height = Number(attr(definition, "rowHeight") ?? 241_300);
      if (flag(definition, "showCaption", true)) drawing.caption = caption; else delete drawing.caption;
      drawing.columnCount = Number.isFinite(columns) && columns >= 1 ? Math.min(columns, 20_000) : 1;
      drawing.itemHeight = Number.isFinite(height) && height > 0 ? height : 241_300;
      drawing.style = slicerStyle(attr(definition, "style") ?? styles.defaultSlicerStyle ?? "SlicerStyleLight1", styles, theme);
      pendingSlicers.push({ drawing, cache });
    }
    const visible = drawings.filter(drawing => drawing.kind !== "slicer" || pendingSlicers.some(pending => pending.drawing === drawing));
    if (visible.length) sheet.drawings = visible;
  }
}

/** Slicer look: custom item styles (x14 slicerStyles) plus the caption/frame parts of the table style of the same name, else a built-in style. */
function slicerStyle(name: string, styles: StyleSheet, theme: Theme): SpreadsheetSlicerStyle {
  const items = styles.slicerStyles.get(name), frame = styles.tableStyles.get(name);
  if (!items && !frame) return builtinSlicerStyle(name, theme);
  const style: SpreadsheetSlicerStyle = {};
  for (const [key, value] of Object.entries(items ?? {})) style[key as keyof SpreadsheetSlicerStyle] = value;
  if (frame?.["wholeTable"]) style.wholeTable = frame["wholeTable"];
  if (frame?.["headerRow"]) style.header = frame["headerRow"];
  return style;
}

/** Office's built-in slicer styles (Light/Other/Dark 1-6), approximated from the theme accents. */
function builtinSlicerStyle(name: string, theme: Theme): SpreadsheetSlicerStyle {
  const match = /^SlicerStyle(Light|Dark|Other)([1-6])$/i.exec(name) ?? ["", "Light", "1"];
  const kind = match[1]!.toLowerCase(), index = Number(match[2]);
  const accent = kind === "other" ? (index === 1 ? "#7f7f7f" : theme.colors[3]!) : theme.colors[3 + index] ?? "#4472c4";
  const text = theme.colors[1] ?? "#000000", edges = (color: string): SpreadsheetBorders => ({ left: { style: "thin", color }, right: { style: "thin", color }, top: { style: "thin", color }, bottom: { style: "thin", color } });
  const solid = (color: string): SpreadsheetFill => ({ type: "pattern", pattern: "solid", foreground: color });
  const noData: SpreadsheetStyle = { font: { color: "#bfbfbf" }, fill: solid("#ffffff"), borders: edges("#e0e0e0") };
  if (kind === "dark") return {
    wholeTable: { fill: solid(applyTint(accent, 0.8)), borders: edges(applyTint(accent, 0.4)) }, header: { font: { bold: true, color: text }, borders: { bottom: { style: "thin", color: accent } } },
    selectedItemWithData: { font: { color: "#ffffff" }, fill: solid(accent), borders: edges(accent) }, selectedItemWithNoData: { font: { color: applyTint(accent, 0.6) }, fill: solid(applyTint(accent, 0.4)), borders: edges(applyTint(accent, 0.4)) },
    unselectedItemWithData: { font: { color: text }, fill: solid(applyTint(accent, 0.6)), borders: edges(applyTint(accent, 0.6)) }, unselectedItemWithNoData: noData,
  };
  return {
    wholeTable: { borders: edges("#bfbfbf") }, header: { font: { bold: true, color: text }, borders: { bottom: { style: "thin", color: kind === "other" ? "#bfbfbf" : applyTint(accent, 0.4) } } },
    selectedItemWithData: { font: { color: "#000000" }, fill: solid(applyTint(accent, 0.6)), borders: edges(applyTint(accent, 0.4)) }, selectedItemWithNoData: { font: { color: "#a6a6a6" }, fill: solid(applyTint(accent, 0.8)), borders: edges(applyTint(accent, 0.6)) },
    unselectedItemWithData: { font: { color: "#000000" }, fill: solid("#ffffff"), borders: edges("#bfbfbf") }, unselectedItemWithNoData: noData,
  };
}

/** Fills slicer buttons from their caches: table slicers read the table column (selection from its filter), pivot slicers the pivot cache items. */
function resolveSlicers(zip: ZipArchive, workbook: XmlNode, relationships: Map<string, Relationship>, pending: PendingSlicer[], tables: Map<number, TableInfo>, locale: string): void {
  const caches = new Map<string, XmlNode>();
  for (const rel of relationships.values()) if (!rel.external && rel.type.endsWith("/slicerCache") && zip.has(rel.target)) { const node = decodeXml(zip, rel.target); caches.set(attr(node, "name") ?? "", node); }
  const pivotCaches = new Map<string, string>();
  const pivotLists = [child(workbook, "pivotCaches"), ...elements(child(workbook, "extLst") ?? workbook, "ext").map(ext => child(ext, "pivotCaches"))];
  for (const list of pivotLists) for (const node of list ? elements(list, "pivotCache") : []) {
    const rel = relationships.get(Object.entries(node.attributes).find(([key]) => key.endsWith(":id"))?.[1] ?? "");
    if (rel && !rel.external && zip.has(rel.target)) pivotCaches.set(attr(node, "cacheId") ?? "", rel.target);
  }
  const collator = new Intl.Collator(locale, { numeric: false, sensitivity: "base" });
  type Item = { text: string; number?: number; selected: boolean; hasData: boolean };
  for (const { drawing, cache: cacheName } of pending) {
    const cache = caches.get(cacheName);
    if (!cache) continue;
    const tableCache = elements(child(cache, "extLst") ?? cache, "ext").map(ext => child(ext, "tableSlicerCache")).find((node): node is XmlNode => !!node);
    const tabular = child(child(cache, "data"), "tabular");
    const settings = tableCache ?? tabular;
    let items: Item[] = [];
    if (tableCache) {
      const table = tables.get(Number(attr(tableCache, "tableId"))), columnIndex = table ? table.columnIds.indexOf(Number(attr(tableCache, "column"))) : -1;
      if (!table || columnIndex < 0) continue;
      const cells = new Map(table.sheet.cells.map(cell => [cell.ref.toUpperCase(), cell]));
      const textAt = (row: number, column: number) => { const cell = cells.get(`${columnName(column)}${row}`); return cell ? (cell.display ?? cell.value) : ""; };
      const passes = (filter: { values: Set<string>; blank: boolean } | undefined, text: string) => !filter || (text === "" ? filter.blank : filter.values.has(text.toLowerCase()));
      const own = table.filters.get(columnIndex), byText = new Map<string, Item>();
      for (let row = table.range.top + table.headerRows; row <= table.range.bottom - table.totalsRows; row++) {
        const text = textAt(row, table.range.left + columnIndex), key = text.toLowerCase();
        // An item has data when its row survives the filters of the table's other columns.
        const visible = [...table.filters].every(([index, filter]) => index === columnIndex || passes(filter, textAt(row, table.range.left + index)));
        const raw = cells.get(`${columnName(table.range.left + columnIndex)}${row}`), number = raw && raw.type !== "s" && raw.type !== "str" && raw.value.trim() !== "" ? Number(raw.value) : NaN;
        const item = byText.get(key) ?? { text: text === "" ? "(blank)" : text, ...(Number.isFinite(number) ? { number } : {}), selected: passes(own, text), hasData: false };
        item.hasData ||= visible;
        byText.set(key, item);
      }
      items = [...byText.values()];
    } else if (tabular) {
      const path = pivotCaches.get(attr(tabular, "pivotCacheId") ?? ""), source = attr(cache, "sourceName");
      if (!path || !source) continue;
      const field = elements(child(decodeXml(zip, path), "cacheFields") ?? cache, "cacheField").find(node => attr(node, "name") === source);
      const shared = (field ? child(field, "sharedItems")?.children ?? [] : []).filter((part): part is XmlNode => typeof part !== "string");
      for (const entry of elements(child(tabular, "items") ?? tabular, "i")) {
        const value = shared[Number(attr(entry, "x"))];
        if (!value) continue;
        const raw = value.localName === "m" ? "" : attr(value, "v") ?? "", number = value.localName === "n" ? Number(raw) : NaN;
        items.push({ text: raw === "" ? "(blank)" : value.localName === "b" ? (raw === "1" ? "TRUE" : "FALSE") : raw, ...(Number.isFinite(number) ? { number } : {}), selected: flag(entry, "s"), hasData: !flag(entry, "nd") });
      }
    } else continue;
    const order = attr(settings, "sortOrder") ?? "ascending", crossFilter = attr(settings, "crossFilter") ?? "showItemsWithDataAtTop";
    if (order !== "natural") {
      // Numbers before text, (blank) last; descending reverses the value order.
      items.sort((a, b) => (a.text === "(blank)" ? 1 : 0) - (b.text === "(blank)" ? 1 : 0) || (a.number !== undefined && b.number !== undefined ? a.number - b.number : a.number !== undefined ? -1 : b.number !== undefined ? 1 : collator.compare(a.text, b.text)) * (order === "descending" ? -1 : 1));
    }
    if (crossFilter === "none") for (const item of items) item.hasData = true;
    if (crossFilter === "showItemsWithDataAtTop") items = [...items.filter(item => item.hasData), ...items.filter(item => !item.hasData)];
    drawing.items = items.map(item => ({ text: item.text, selected: item.selected, hasData: item.hasData }));
    drawing.filtered = items.some(item => !item.selected) && items.some(item => item.selected);
    if (!items.some(item => item.selected)) for (const item of drawing.items) item.selected = true;
  }
}

/** Evaluates conditional formatting into per-cell overlays (Excel applies all matching rules; higher priority wins per property). */
function evaluateConditionalFormats(root: XmlNode, sheet: Worksheet, styles: StyleSheet, theme: Theme, palette: string[], value: (row: number, column: number) => CellValue, context: (row: number, column: number) => FormulaContext, locale: string, date1904: boolean): Record<string, SpreadsheetConditional> {
  type Rule = { node: XmlNode; ranges: SpreadsheetRange[]; priority: number; dxf?: SpreadsheetStyle; ext?: XmlNode; formulas: string[] };
  const rules: Rule[] = [];
  const extById = new Map<string, XmlNode>(), extRules: Rule[] = [];
  const extList = child(root, "extLst");
  for (const ext of elements(extList ?? root, "ext")) for (const block of elements(child(ext, "conditionalFormattings") ?? ext, "conditionalFormatting")) {
    const ranges = (textContent(child(block, "sqref") ?? block) || "").trim().split(/\s+/).map(parseRange).filter((range): range is SpreadsheetRange => !!range);
    for (const rule of elements(block, "cfRule")) {
      const id = attr(rule, "id");
      if (id) extById.set(id, rule);
      if (attr(rule, "type") === "dataBar" && id) continue; // merged into the matching main rule
      const dxfNode = child(rule, "dxf");
      extRules.push({ node: rule, ranges, priority: Number(attr(rule, "priority") ?? 1000), ...(dxfNode ? { dxf: readDxf(dxfNode, theme, palette, locale) } : {}), formulas: elements(rule, "f").map(textContent) });
    }
  }
  for (const block of elements(root, "conditionalFormatting")) {
    const ranges = (attr(block, "sqref") ?? "").trim().split(/\s+/).map(parseRange).filter((range): range is SpreadsheetRange => !!range);
    for (const rule of elements(block, "cfRule")) {
      const dxfId = attr(rule, "dxfId"), extId = textContent(child(child(child(rule, "extLst"), "ext"), "id") ?? { name: "", localName: "", attributes: {}, children: [] }).trim();
      const ext = extId ? extById.get(extId) : undefined;
      rules.push({ node: rule, ranges, priority: Number(attr(rule, "priority") ?? 1000), ...(dxfId !== undefined && styles.dxfs[Number(dxfId)] ? { dxf: styles.dxfs[Number(dxfId)]! } : {}), ...(ext ? { ext } : {}), formulas: elements(rule, "formula").map(textContent) });
    }
  }
  const extIds = new Set(rules.map(rule => rule.ext && attr(rule.ext, "id")).filter(Boolean));
  for (const rule of extRules) if (!extIds.has(attr(rule.node, "id"))) rules.push(rule);
  if (!rules.length) return {};
  rules.sort((a, b) => a.priority - b.priority);
  const limitRow = sheet.maxRow + 60, limitColumn = sheet.maxColumn + 30;
  const result: Record<string, SpreadsheetConditional> = {}, stopped = new Set<string>();
  let budget = MAX_CF_CELLS;
  const color = (node: XmlNode | undefined) => readColor(node, theme, palette);
  for (const rule of rules) {
    const type = attr(rule.node, "type") ?? "";
    const cells: Array<{ row: number; column: number }> = [];
    for (const range of rule.ranges) for (let row = range.top; row <= Math.min(range.bottom, limitRow); row++) for (let column = range.left; column <= Math.min(range.right, limitColumn); column++) {
      if (--budget < 0) break;
      cells.push({ row, column });
    }
    if (!cells.length) continue;
    const origin = { row: Math.min(...rule.ranges.map(range => range.top)), column: Math.min(...rule.ranges.map(range => range.left)) };
    if (rule.ranges[0]) { origin.row = rule.ranges[0].top; origin.column = rule.ranges[0].left; }
    const numbers = () => cells.map(cell => value(cell.row, cell.column)).filter((entry): entry is number => typeof entry === "number");
    const threshold = (cfvo: XmlNode | undefined, values: number[], fallback: "min" | "max"): number => {
      const kind = attr(cfvo, "type") ?? fallback, raw = attr(cfvo, "val") ?? textContent(child(cfvo, "f") ?? cfvo ?? { name: "", localName: "", attributes: {}, children: [] });
      const min = values.length ? Math.min(...values) : 0, max = values.length ? Math.max(...values) : 0;
      if (kind === "min") return min;
      if (kind === "max") return max;
      if (kind === "autoMin") return Math.min(0, min);
      if (kind === "autoMax") return Math.max(0, max);
      let number = Number(raw);
      if (kind === "formula" || !Number.isFinite(number)) { const evaluated = evaluateFormula(raw, context(origin.row, origin.column), origin); number = typeof evaluated === "number" ? evaluated : 0; }
      if (kind === "percent") return min + (max - min) * number / 100;
      if (kind === "percentile") { const sorted = [...values].sort((a, b) => a - b); if (!sorted.length) return 0; const rank = (sorted.length - 1) * number / 100, low = Math.floor(rank); return sorted[low]! + (sorted[Math.min(low + 1, sorted.length - 1)]! - sorted[low]!) * (rank - low); }
      return number;
    };
    const apply = (row: number, column: number, update: (entry: SpreadsheetConditional) => void) => {
      const key = `${columnName(column)}${row}`;
      if (stopped.has(key)) return;
      const entry = result[key] ??= {};
      update(entry);
      if (flag(rule.node, "stopIfTrue")) stopped.add(key);
    };
    const applyDxf = (row: number, column: number) => apply(row, column, entry => {
      if (!rule.dxf) return;
      entry.style = mergeUnder(entry.style, rule.dxf);
      if (rule.dxf.numberFormat && !entry.display) {
        const current = value(row, column);
        if (typeof current === "number") entry.display = formatValue(current, rule.dxf.numberFormat, { locale, date1904 }).text;
      }
    });
    if (type === "colorScale") {
      const scale = child(rule.node, "colorScale"), cfvos = elements(scale ?? rule.node, "cfvo"), colors = elements(scale ?? rule.node, "color").map(node => color(node) ?? "#ffffff");
      const values = numbers(), stops = cfvos.map((cfvo, index) => threshold(cfvo, values, index === 0 ? "min" : "max"));
      for (const cell of cells) {
        const current = value(cell.row, cell.column);
        if (typeof current !== "number") continue;
        apply(cell.row, cell.column, entry => { if (!entry.style?.fill) entry.style = { ...entry.style, fill: { type: "pattern", pattern: "solid", foreground: interpolateScale(current, stops, colors) } }; });
      }
      continue;
    }
    if (type === "dataBar") {
      const bar = child(rule.node, "dataBar"), ext = child(rule.ext, "dataBar"), cfvos = elements(bar ?? rule.node, "cfvo");
      const values = numbers();
      const extCfvos = ext ? elements(ext, "cfvo") : [];
      const low = threshold(extCfvos[0] ?? cfvos[0], values, "min"), high = threshold(extCfvos[1] ?? cfvos[1], values, "max");
      const minLength = Number(attr(ext, "minLength") ?? attr(bar, "minLength") ?? (ext ? 0 : 10)) / 100, maxLength = Number(attr(ext, "maxLength") ?? attr(bar, "maxLength") ?? (ext ? 100 : 90)) / 100;
      const barColor = color(child(bar, "color")) ?? "#638ec6", gradient = attr(ext, "gradient") !== "0", border = ext && attr(ext, "border") === "1" ? color(child(ext, "borderColor")) ?? barColor : undefined;
      const negativeColor = color(child(ext, "negativeFillColor")) ?? "#ff0000", axisNone = attr(ext, "axisPosition") === "none";
      const hide = attr(bar, "showValue") === "0";
      for (const cell of cells) {
        const current = value(cell.row, cell.column);
        if (typeof current !== "number") continue;
        apply(cell.row, cell.column, entry => {
          if (entry.dataBar) return;
          if (!axisNone && low < 0 && high > 0) {
            const axis = -low / (high - low), length = Math.min(1, Math.abs(current) / (high - low));
            entry.dataBar = current >= 0 ? { color: barColor, start: axis, end: axis + length, gradient, axis, ...(border ? { border } : {}) } : { color: negativeColor, start: axis - length, end: axis, gradient, axis, negative: true };
          } else {
            const fraction = high === low ? 1 : Math.max(0, Math.min(1, (current - low) / (high - low)));
            entry.dataBar = { color: barColor, start: 0, end: minLength + (maxLength - minLength) * fraction, gradient, ...(border ? { border } : {}) };
          }
          if (hide) entry.hideValue = true;
        });
      }
      continue;
    }
    if (type === "iconSet") {
      const set = child(rule.ext, "iconSet") ?? child(rule.node, "iconSet"), cfvos = elements(set ?? rule.node, "cfvo");
      const iconSet = attr(set, "iconSet") ?? "3TrafficLights1", reverse = attr(set, "reverse") === "1", hide = attr(set, "showValue") === "0";
      const values = numbers(), limits = cfvos.map(cfvo => ({ value: threshold(cfvo, values, "min"), gte: attr(cfvo, "gte") !== "0" }));
      const custom = elements(set ?? rule.node, "cfIcon");
      for (const cell of cells) {
        const current = value(cell.row, cell.column);
        if (typeof current !== "number") continue;
        let index = 0;
        limits.forEach((limit, position) => { if (position > 0 && (limit.gte ? current >= limit.value : current > limit.value)) index = position; });
        const iconIndex = reverse ? limits.length - 1 - index : index;
        const customIcon = custom[index];
        apply(cell.row, cell.column, entry => {
          if (entry.icon) return;
          entry.icon = customIcon ? { set: attr(customIcon, "iconSet") ?? iconSet, index: Number(attr(customIcon, "iconId") ?? iconIndex) } : { set: iconSet, index: iconIndex };
          if (hide) entry.hideValue = true;
        });
      }
      continue;
    }
    // Rules with a differential format: decide per cell whether the rule matches.
    const values = cells.map(cell => value(cell.row, cell.column));
    let matcher: (current: CellValue, row: number, column: number) => boolean;
    const formulaValue = (index: number, row: number, column: number) => { const formula = rule.formulas[index]; return formula === undefined ? null : evaluateFormula(formula, context(row, column), origin); };
    const scalar = (input: unknown): CellValue => typeof input === "object" && input !== null && !(input instanceof FormulaError) ? null : input as CellValue;
    const numericValues = values.filter((entry): entry is number => typeof entry === "number");
    if (type === "expression") matcher = (_current, row, column) => isTruthy(formulaValue(0, row, column));
    else if (type === "cellIs") {
      const operator = attr(rule.node, "operator") ?? "equal";
      matcher = (current, row, column) => {
        if (current === null && operator !== "equal" && operator !== "notEqual") current = 0;
        const a = scalar(formulaValue(0, row, column)), b = scalar(formulaValue(1, row, column));
        const cmp = (x: CellValue, y: CellValue) => compareValues(x, y);
        switch (operator) {
          case "lessThan": return cmp(current, a) < 0;
          case "lessThanOrEqual": return cmp(current, a) <= 0;
          case "greaterThan": return cmp(current, a) > 0;
          case "greaterThanOrEqual": return cmp(current, a) >= 0;
          case "notEqual": return cmp(current, a) !== 0;
          case "between": return cmp(current, minValue(a, b)) >= 0 && cmp(current, maxValue(a, b)) <= 0;
          case "notBetween": return cmp(current, minValue(a, b)) < 0 || cmp(current, maxValue(a, b)) > 0;
          default: return cmp(current, a) === 0;
        }
      };
    } else if (type === "containsText" || type === "notContainsText" || type === "beginsWith" || type === "endsWith") {
      const needle = (attr(rule.node, "text") ?? "").toLowerCase();
      matcher = current => {
        const text = current === null ? "" : current instanceof FormulaError ? current.code : String(current).toLowerCase();
        return type === "containsText" ? text.includes(needle) : type === "notContainsText" ? !text.includes(needle) : type === "beginsWith" ? text.startsWith(needle) : text.endsWith(needle);
      };
    } else if (type === "containsBlanks") matcher = current => current === null || typeof current === "string" && current.trim() === "";
    else if (type === "notContainsBlanks") matcher = current => !(current === null || typeof current === "string" && current.trim() === "");
    else if (type === "containsErrors") matcher = current => current instanceof FormulaError;
    else if (type === "notContainsErrors") matcher = current => !(current instanceof FormulaError);
    else if (type === "duplicateValues" || type === "uniqueValues") {
      const counts = new Map<string, number>();
      for (const entry of values) if (entry !== null && entry !== "") { const key = typeof entry === "string" ? `s${entry.toLowerCase()}` : `v${String(entry)}`; counts.set(key, (counts.get(key) ?? 0) + 1); }
      matcher = current => { if (current === null || current === "") return false; const count = counts.get(typeof current === "string" ? `s${current.toLowerCase()}` : `v${String(current)}`) ?? 0; return type === "duplicateValues" ? count > 1 : count === 1; };
    } else if (type === "top10") {
      const rank = Number(attr(rule.node, "rank") ?? 10), percent = flag(rule.node, "percent"), bottom = flag(rule.node, "bottom");
      const sorted = [...numericValues].sort((a, b) => bottom ? a - b : b - a), count = Math.max(1, percent ? Math.floor(sorted.length * rank / 100) : rank);
      const limit = sorted[Math.min(count, sorted.length) - 1];
      matcher = current => typeof current === "number" && limit !== undefined && (bottom ? current <= limit : current >= limit);
    } else if (type === "aboveAverage") {
      const above = attr(rule.node, "aboveAverage") !== "0", equal = flag(rule.node, "equalAverage"), deviations = Number(attr(rule.node, "stdDev") ?? 0);
      const mean = numericValues.reduce((sum, entry) => sum + entry, 0) / Math.max(1, numericValues.length);
      const deviation = Math.sqrt(numericValues.reduce((sum, entry) => sum + (entry - mean) ** 2, 0) / Math.max(1, numericValues.length - 1));
      const limit = mean + (above ? 1 : -1) * deviations * deviation;
      matcher = current => typeof current === "number" && (above ? (equal ? current >= limit : current > limit) : (equal ? current <= limit : current < limit));
    } else if (type === "timePeriod") {
      const today = context(origin.row, origin.column).today ?? todaySerial(), period = attr(rule.node, "timePeriod") ?? "today";
      const weekday = (today + 6) % 7; // 0 = Sunday for Excel serials
      const ranges: Record<string, [number, number]> = {
        today: [today, today], yesterday: [today - 1, today - 1], tomorrow: [today + 1, today + 1], last7Days: [today - 6, today],
        thisWeek: [today - weekday, today - weekday + 6], lastWeek: [today - weekday - 7, today - weekday - 1], nextWeek: [today - weekday + 7, today - weekday + 13],
      };
      const monthRange = (offset: number): [number, number] => { const base = new Date(Date.UTC(1899, 11, 30 + today)); const start = new Date(Date.UTC(base.getUTCFullYear(), base.getUTCMonth() + offset, 1)), end = new Date(Date.UTC(base.getUTCFullYear(), base.getUTCMonth() + offset + 1, 0)); return [dateToSerial(start.getUTCFullYear(), start.getUTCMonth(), start.getUTCDate()), dateToSerial(end.getUTCFullYear(), end.getUTCMonth(), end.getUTCDate())]; };
      const [from, to] = ranges[period] ?? (period === "thisMonth" ? monthRange(0) : period === "lastMonth" ? monthRange(-1) : period === "nextMonth" ? monthRange(1) : [today, today]);
      matcher = current => typeof current === "number" && Math.floor(current) >= from && Math.floor(current) <= to;
    } else if (rule.formulas.length) matcher = (_current, row, column) => isTruthy(formulaValue(0, row, column));
    else continue;
    cells.forEach((cell, index) => {
      try { if (matcher(values[index]!, cell.row, cell.column)) applyDxf(cell.row, cell.column); }
      catch { /* malformed rule: ignore for this cell */ }
    });
  }
  return result;
}

/** Sparkline groups (Excel 2010 extension): one in-cell chart per target cell, with the group's type, colors and axis scaling. */
function readSparklines(root: XmlNode, sheetName: string, value: (sheet: string, row: number, column: number) => CellValue, theme: Theme, palette: string[]): SpreadsheetSparkline[] {
  const result: SpreadsheetSparkline[] = [];
  const color = (node: XmlNode | undefined) => readColor(node, theme, palette);
  for (const ext of elements(child(root, "extLst") ?? root, "ext")) for (const group of elements(child(ext, "sparklineGroups") ?? ext, "sparklineGroup")) {
    const kind = attr(group, "type"), type = kind === "column" || kind === "stacked" ? kind : "line";
    const empty = attr(group, "displayEmptyCellsAs"), emptyCells = empty === "gap" || empty === "span" ? empty : "zero";
    const weight = Number(attr(group, "lineWeight") ?? 0.75), shown = (name: string, colorName: string) => flag(group, name) ? color(child(group, colorName)) : undefined;
    const optional = { negativeColor: shown("negative", "colorNegative"), axisColor: shown("displayXAxis", "colorAxis"), markersColor: type === "line" ? shown("markers", "colorMarkers") : undefined, highColor: shown("high", "colorHigh"), lowColor: shown("low", "colorLow"), firstColor: shown("first", "colorFirst"), lastColor: shown("last", "colorLast") };
    const members: SpreadsheetSparkline[] = [];
    for (const sparkline of elements(child(group, "sparklines") ?? group, "sparkline")) {
      const ref = textContent(child(sparkline, "sqref") ?? sparkline).trim(), formula = textContent(child(sparkline, "f") ?? { name: "", localName: "", attributes: {}, children: [] }).trim();
      const match = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(formula), range = parseRange(match ? match[3]! : formula);
      if (!range || !parseCellReference(ref).row) continue;
      const source = match ? (match[1]?.replace(/''/g, "'") ?? match[2]!) : sheetName, values: Array<number | null> = [];
      for (let row = range.top; row <= Math.min(range.bottom, range.top + 1000); row++) for (let column = range.left; column <= Math.min(range.right, range.left + 1000); column++) {
        const cell = value(source, row, column);
        values.push(typeof cell === "number" ? cell : typeof cell === "boolean" ? Number(cell) : null);
      }
      const entry: SpreadsheetSparkline = { ref, type, values, color: color(child(group, "colorSeries")) ?? "#376092", lineWeight: Number.isFinite(weight) && weight > 0 ? weight : 0.75, emptyCells };
      for (const [key, option] of Object.entries(optional)) if (option) (entry as Record<string, unknown>)[key] = option;
      if (flag(group, "rightToLeft")) entry.rightToLeft = true;
      members.push(entry);
    }
    // Axis bounds shared by the group or fixed by the author; "individual" scaling leaves them per sparkline.
    const numbers = members.flatMap(member => member.values.filter((entry): entry is number => entry !== null));
    for (const [bound, typeName, manualName] of [["min", "minAxisType", "manualMin"], ["max", "maxAxisType", "manualMax"]] as const) {
      const axisType = attr(group, typeName), manual = Number(attr(group, manualName) ?? 0);
      const fixed = axisType === "custom" ? manual : axisType === "group" && numbers.length ? (bound === "min" ? Math.min(...numbers) : Math.max(...numbers)) : undefined;
      if (fixed !== undefined && Number.isFinite(fixed)) for (const member of members) member[bound] = fixed;
    }
    result.push(...members);
  }
  return result;
}

function minValue(a: CellValue, b: CellValue): CellValue { return compareValues(a, b) <= 0 ? a : b; }
function maxValue(a: CellValue, b: CellValue): CellValue { return compareValues(a, b) <= 0 ? b : a; }
function compareValues(a: CellValue, b: CellValue): number {
  if (a instanceof FormulaError || b instanceof FormulaError) return NaN;
  if (a === null) a = typeof b === "string" ? "" : 0;
  if (b === null) b = typeof a === "string" ? "" : 0;
  const rank = (value: CellValue) => typeof value === "number" ? 0 : typeof value === "string" ? 1 : 2;
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === "string" && typeof b === "string") return a.toLowerCase().localeCompare(b.toLowerCase());
  return Number(a) - Number(b);
}

/** Merges `lower` (lower priority) underneath `upper`: properties already set by a higher-priority rule win. */
function mergeUnder(upper: SpreadsheetStyle | undefined, lower: SpreadsheetStyle): SpreadsheetStyle {
  if (!upper) return { ...lower };
  const merged: SpreadsheetStyle = { ...upper };
  if (lower.font) merged.font = { ...lower.font, ...upper.font };
  if (lower.fill && !upper.fill) merged.fill = lower.fill;
  if (lower.borders) merged.borders = { ...lower.borders, ...upper.borders };
  if (lower.numberFormat && !upper.numberFormat) merged.numberFormat = lower.numberFormat;
  return merged;
}

function interpolateScale(value: number, stops: number[], colors: string[]): string {
  if (!stops.length) return colors[0] ?? "#ffffff";
  if (value <= stops[0]!) return colors[0]!;
  for (let index = 1; index < stops.length; index++) {
    if (value <= stops[index]!) {
      const span = stops[index]! - stops[index - 1]!, t = span ? (value - stops[index - 1]!) / span : 1;
      const a = hexToRgb(colors[index - 1]!), b = hexToRgb(colors[index] ?? colors[index - 1]!);
      return rgbToHex(a.map((channel, i) => Math.round(channel + (b[i]! - channel) * t)) as [number, number, number]);
    }
  }
  return colors[Math.min(stops.length, colors.length) - 1]!;
}

// ---------------------------------------------------------------------------------------------------------------------
// Styles

function readNumberFormats(root: XmlNode | undefined): Map<number, string> {
  const formats = new Map<number, string>();
  for (const node of elements(child(root, "numFmts") ?? { name: "", localName: "", attributes: {}, children: [] }, "numFmt")) {
    const id = Number(attr(node, "numFmtId")), code = attr(node, "formatCode");
    if (Number.isFinite(id) && code !== undefined) formats.set(id, code);
  }
  return formats;
}

function readStyles(root: XmlNode | undefined, theme: Theme, palette: string[], locale: string): StyleSheet {
  const empty = { name: "", localName: "", attributes: {}, children: [] } as XmlNode;
  const custom = readNumberFormats(root);
  const fonts = elements(child(root, "fonts") ?? empty, "font").map(node => readFont(node, theme, palette));
  const fills = elements(child(root, "fills") ?? empty, "fill").map(node => readFill(node, theme, palette, false));
  const borders = elements(child(root, "borders") ?? empty, "border").map(node => readBorders(node, theme, palette));
  const cellStyles: SpreadsheetStyle[] = [], fontIds: number[] = [];
  const xfs = elements(child(root, "cellXfs") ?? empty, "xf");
  for (const xf of xfs.length ? xfs : [empty]) {
    const style: SpreadsheetStyle = {};
    const fontId = Number(attr(xf, "fontId") ?? 0), fillId = Number(attr(xf, "fillId") ?? 0), borderId = Number(attr(xf, "borderId") ?? 0), numFmtId = Number(attr(xf, "numFmtId") ?? 0);
    fontIds.push(fontId);
    if (fonts[fontId]) style.font = fonts[fontId];
    if (fills[fillId]) style.fill = fills[fillId]!;
    if (borders[borderId] && Object.keys(borders[borderId]!).length) style.borders = borders[borderId]!;
    const code = custom.get(numFmtId) ?? builtinNumberFormat(numFmtId, locale);
    if (code && code !== "General") style.numberFormat = code;
    const alignment = readAlignment(child(xf, "alignment"));
    if (alignment) style.alignment = alignment;
    cellStyles.push(style);
  }
  const dxfs = elements(child(root, "dxfs") ?? empty, "dxf").map(node => readDxf(node, theme, palette, locale, custom));
  const tableStyles = new Map<string, Record<string, SpreadsheetStyle & { size?: number }>>();
  for (const node of elements(child(root, "tableStyles") ?? empty, "tableStyle")) {
    const name = attr(node, "name");
    if (!name) continue;
    const entry: Record<string, SpreadsheetStyle & { size?: number }> = {};
    for (const element of elements(node, "tableStyleElement")) {
      const type = attr(element, "type"), dxf = dxfs[Number(attr(element, "dxfId"))];
      if (type && dxf) entry[type] = { ...dxf, ...(attr(element, "size") ? { size: Number(attr(element, "size")) } : {}) };
    }
    tableStyles.set(name, entry);
  }
  // Slicer item styles (Excel 2010 extension) index their own differential format list, kept in a sibling extension.
  const extensions = elements(child(root, "extLst") ?? empty, "ext");
  const extensionDxfs = elements(extensions.map(ext => child(ext, "dxfs")).find((node): node is XmlNode => !!node) ?? empty, "dxf").map(node => readDxf(node, theme, palette, locale, custom));
  const slicerStyles = new Map<string, Record<string, SpreadsheetStyle>>();
  let defaultSlicerStyle: string | undefined;
  for (const list of extensions.map(ext => child(ext, "slicerStyles")).filter((node): node is XmlNode => !!node)) {
    defaultSlicerStyle = attr(list, "defaultSlicerStyle") ?? defaultSlicerStyle;
    for (const node of elements(list, "slicerStyle")) {
      const entry: Record<string, SpreadsheetStyle> = {};
      for (const element of elements(child(node, "slicerStyleElements") ?? node, "slicerStyleElement")) {
        const type = attr(element, "type"), dxf = extensionDxfs[Number(attr(element, "dxfId"))];
        if (type && dxf) entry[type] = dxf;
      }
      slicerStyles.set(attr(node, "name") ?? "", entry);
    }
  }
  return { cellStyles, dxfs, tableStyles, slicerStyles, ...(defaultSlicerStyle ? { defaultSlicerStyle } : {}), defaultFont: fonts[0] ?? { family: "Calibri", size: 11 }, fontIds };
}

function readFont(node: XmlNode, theme: Theme, palette: string[], color?: (node: XmlNode | undefined) => string | undefined): SpreadsheetFont {
  const font: SpreadsheetFont = {};
  const scheme = attr(child(node, "scheme"), "val");
  const family = attr(child(node, "name"), "val") ?? attr(child(node, "rFont"), "val") ?? (scheme === "major" ? theme.majorFont : scheme === "minor" ? theme.minorFont : undefined);
  if (family) font.family = family;
  const size = Number(attr(child(node, "sz"), "val"));
  if (Number.isFinite(size) && size > 0) font.size = size;
  if (on(child(node, "b"))) font.bold = true;
  if (on(child(node, "i"))) font.italic = true;
  if (on(child(node, "strike"))) font.strike = true;
  const underline = child(node, "u");
  if (underline) { const kind = attr(underline, "val") ?? "single"; if (kind !== "none") font.underline = kind as NonNullable<SpreadsheetFont["underline"]>; }
  const vertAlign = attr(child(node, "vertAlign"), "val");
  if (vertAlign === "superscript" || vertAlign === "subscript") font.vertAlign = vertAlign;
  const fontColor = color ? color(child(node, "color")) : readColor(child(node, "color"), theme, palette, true);
  if (fontColor) font.color = fontColor;
  return font;
}

function readFill(node: XmlNode, theme: Theme, palette: string[], differential: boolean): SpreadsheetFill | undefined {
  const pattern = child(node, "patternFill");
  if (pattern) {
    const kind = attr(pattern, "patternType") ?? (differential ? "solid" : "none");
    if (kind === "none") return undefined;
    const fg = readColor(child(pattern, "fgColor"), theme, palette), bg = readColor(child(pattern, "bgColor"), theme, palette);
    // Differential (conditional) solid fills take their color from bgColor; cell fills from fgColor.
    if (kind === "solid") { const solid = differential ? bg ?? fg : fg ?? bg ?? "#000000"; return solid ? { type: "pattern", pattern: "solid", foreground: solid } : undefined; }
    return { type: "pattern", pattern: kind, foreground: fg ?? "#000000", background: bg ?? "#ffffff" };
  }
  const gradient = child(node, "gradientFill");
  if (gradient) {
    const stops = elements(gradient, "stop").map(stop => ({ position: Number(attr(stop, "position") ?? 0), color: readColor(child(stop, "color"), theme, palette) ?? "#ffffff" }));
    const number = (name: string) => { const value = Number(attr(gradient, name)); return Number.isFinite(value) ? value : undefined; };
    const fill: SpreadsheetFill = { type: "gradient", gradient: attr(gradient, "type") === "path" ? "path" : "linear", degree: number("degree") ?? 0, stops };
    for (const side of ["left", "right", "top", "bottom"] as const) { const value = number(side); if (value !== undefined) fill[side] = value; }
    return stops.length ? fill : undefined;
  }
  return undefined;
}

function readBorders(node: XmlNode, theme: Theme, palette: string[]): SpreadsheetBorders {
  const borders: SpreadsheetBorders = {};
  for (const [name, side] of [["left", "left"], ["start", "left"], ["right", "right"], ["end", "right"], ["top", "top"], ["bottom", "bottom"], ["diagonal", "diagonal"]] as const) {
    const edge = child(node, name), style = attr(edge, "style");
    if (!edge || !style || style === "none" || borders[side]) continue;
    borders[side] = { style, color: readColor(child(edge, "color"), theme, palette, true) ?? "#000000" };
  }
  for (const [name, side] of [["vertical", "insideV"], ["horizontal", "insideH"]] as const) {
    const edge = child(node, name), style = attr(edge, "style");
    if (edge && style && style !== "none") (borders as Record<string, SpreadsheetBorderEdge>)[side] = { style, color: readColor(child(edge, "color"), theme, palette, true) ?? "#000000" };
  }
  if (borders.diagonal) { if (flag(node, "diagonalUp")) borders.diagonalUp = true; if (flag(node, "diagonalDown")) borders.diagonalDown = true; }
  return borders;
}

function readAlignment(node: XmlNode | undefined): SpreadsheetStyle["alignment"] {
  if (!node) return undefined;
  const alignment: NonNullable<SpreadsheetStyle["alignment"]> = {};
  const horizontal = attr(node, "horizontal"), vertical = attr(node, "vertical");
  if (horizontal && horizontal !== "general") alignment.horizontal = horizontal as NonNullable<typeof alignment.horizontal>;
  if (vertical) alignment.vertical = vertical as NonNullable<typeof alignment.vertical>;
  if (flag(node, "wrapText")) alignment.wrapText = true;
  if (flag(node, "shrinkToFit")) alignment.shrinkToFit = true;
  const indent = Number(attr(node, "indent")), rotation = Number(attr(node, "textRotation"));
  if (indent > 0) alignment.indent = indent;
  if (rotation > 0 && rotation <= 255) alignment.textRotation = rotation;
  return Object.keys(alignment).length ? alignment : undefined;
}

function readDxf(node: XmlNode, theme: Theme, palette: string[], locale: string, formats?: Map<number, string>): SpreadsheetStyle {
  const style: SpreadsheetStyle = {};
  const font = child(node, "font");
  if (font) { const parsed = readFont(font, theme, palette); delete parsed.family; delete parsed.size; if (Object.keys(parsed).length) style.font = parsed; }
  const fill = child(node, "fill");
  if (fill) { const parsed = readFill(fill, theme, palette, true); if (parsed) style.fill = parsed; }
  const border = child(node, "border");
  if (border) { const parsed = readBorders(border, theme, palette); if (Object.keys(parsed).length) style.borders = parsed; }
  const numFmt = child(node, "numFmt");
  if (numFmt) { const code = attr(numFmt, "formatCode") ?? formats?.get(Number(attr(numFmt, "numFmtId"))) ?? builtinNumberFormat(Number(attr(numFmt, "numFmtId")), locale); if (code) style.numberFormat = code; }
  const alignment = readAlignment(child(node, "alignment"));
  if (alignment) style.alignment = alignment;
  return style;
}

function readRichText(node: XmlNode, theme?: Theme, palette?: string[], color?: (node: XmlNode | undefined) => string | undefined): { text: string; runs?: SpreadsheetRun[] } {
  const runs: SpreadsheetRun[] = [];
  let plain = "";
  for (const part of node.children) {
    if (typeof part === "string") continue;
    if (part.localName === "t") plain += decodeEscapes(textContent(part));
    else if (part.localName === "r") {
      const text = decodeEscapes(elements(part, "t").map(textContent).join("")), properties = child(part, "rPr");
      runs.push(properties ? { text, font: readFont(properties, theme ?? { colors: defaultThemeColors }, palette ?? defaultPalette, color) } : { text });
    }
  }
  const text = plain + runs.map(run => run.text).join("");
  return runs.length && runs.some(run => run.font) ? { text, runs: plain ? [{ text: plain }, ...runs] : runs } : { text };
}

/** Decodes Excel's _xHHHH_ escapes for characters that cannot appear in XML. */
function decodeEscapes(text: string): string { return text.replace(/_x([0-9A-Fa-f]{4})_/g, (_all, hex: string) => String.fromCharCode(Number.parseInt(hex, 16))); }

// ---------------------------------------------------------------------------------------------------------------------
// Colors

const defaultThemeColors = ["#ffffff", "#000000", "#e7e6e6", "#44546a", "#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47", "#0563c1", "#954f72"];
const defaultPalette = ["#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff", ...numberFormatColors];

function readIndexedPalette(root: XmlNode | undefined): string[] {
  const custom = elements(child(child(root, "colors"), "indexedColors") ?? { name: "", localName: "", attributes: {}, children: [] }, "rgbColor").map(node => argbToHex(attr(node, "rgb") ?? "")).filter((value): value is string => !!value);
  return custom.length ? custom : defaultPalette;
}

function readTheme(root: XmlNode | undefined): Theme {
  if (!root) return { colors: defaultThemeColors };
  const elementsNode = child(root, "themeElements") ?? root, scheme = child(elementsNode, "clrScheme");
  const read = (name: string) => {
    const entry = child(scheme, name), color = entry?.children.find((part): part is XmlNode => typeof part !== "string");
    const value = attr(color, "lastClr") ?? attr(color, "val");
    return value && /^[0-9a-f]{6}$/i.test(value) ? `#${value.toLowerCase()}` : undefined;
  };
  // SpreadsheetML theme indexes swap light/dark: 0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2.
  const colors = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"].map((name, index) => read(name) ?? defaultThemeColors[index]!);
  const fonts = child(elementsNode, "fontScheme");
  const latin = (name: string) => attr(child(child(fonts, name), "latin"), "typeface") || undefined;
  const major = latin("majorFont"), minor = latin("minorFont");
  return { colors, ...(major ? { majorFont: major } : {}), ...(minor ? { minorFont: minor } : {}) };
}

function readColor(node: XmlNode | undefined, theme: Theme, palette: string[], isForeground = false): string | undefined {
  if (!node) return undefined;
  let base: string | undefined;
  const rgb = attr(node, "rgb"), themeIndex = attr(node, "theme"), indexed = attr(node, "indexed");
  if (rgb) base = argbToHex(rgb);
  else if (themeIndex !== undefined) base = theme.colors[Number(themeIndex)];
  else if (indexed !== undefined) { const index = Number(indexed); base = index === 64 ? (isForeground ? "#000000" : "#ffffff") : index === 65 ? "#ffffff" : palette[index] ?? defaultPalette[index]; }
  else if (flag(node, "auto")) base = isForeground ? "#000000" : "#ffffff";
  if (!base) return undefined;
  const tint = Number(attr(node, "tint") ?? 0);
  return tint ? applyTint(base, tint) : base;
}

function argbToHex(value: string): string | undefined {
  const hex = value.replace(/^#/, "");
  if (/^[0-9a-f]{8}$/i.test(hex)) return `#${hex.slice(2).toLowerCase()}`;
  if (/^[0-9a-f]{6}$/i.test(hex)) return `#${hex.toLowerCase()}`;
  return undefined;
}

/** Excel tint: adjusts HSL luminance towards white (positive) or black (negative). */
export function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [h, s, l] = rgbToHsl(hexToRgb(hex));
  const luminance = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  return rgbToHex(hslToRgb(h, s, Math.max(0, Math.min(1, luminance))));
}
function hexToRgb(hex: string): [number, number, number] { const value = Number.parseInt(hex.slice(1, 7), 16); return [(value >> 16) & 255, (value >> 8) & 255, value & 255]; }
function rgbToHex(rgb: [number, number, number]): string { return `#${rgb.map(channel => Math.round(Math.max(0, Math.min(255, channel))).toString(16).padStart(2, "0")).join("")}`; }
function rgbToHsl([r, g, b]: [number, number, number]): [number, number, number] {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min, s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}
function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
  const hue = (t: number) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  return [hue(h + 1 / 3) * 255, hue(h) * 255, hue(h - 1 / 3) * 255];
}

// ---------------------------------------------------------------------------------------------------------------------
// Drawings

function readDrawings(zip: ZipArchive, path: string, themeRoot: XmlNode | undefined): SpreadsheetDrawing[] {
  const root = decodeXml(zip, path), rels = readRelationships(zip, path), drawings: SpreadsheetDrawing[] = [];
  const marker = (node: XmlNode | undefined) => node ? { column: Number(textContent(child(node, "col") ?? node)) || 0, row: Number(textContent(child(node, "row") ?? node)) || 0, columnOffset: Number(textContent(child(node, "colOff") ?? node)) || 0, rowOffset: Number(textContent(child(node, "rowOff") ?? node)) || 0 } : undefined;
  const image = (embed: string) => {
    const rel = rels.get(embed);
    if (!rel || rel.external || !zip.has(rel.target)) return undefined;
    // Browsers cannot show TIFF; convert it to PNG.
    if (/\.tiff?$/i.test(rel.target)) { const png = tiffToPng(zip.read(rel.target)); return png ? { bytes: png, mimeType: "image/png" } : undefined; }
    const mimeType = imageMimeType(rel.target);
    return mimeType ? { bytes: zip.read(rel.target), mimeType } : undefined;
  };
  for (const anchorNode of root.children) {
    if (typeof anchorNode === "string" || !/Anchor$/.test(anchorNode.localName)) continue;
    const from = marker(child(anchorNode, "from")), to = marker(child(anchorNode, "to")), ext = child(anchorNode, "ext"), pos = child(anchorNode, "pos");
    const anchor: SpreadsheetAnchor = { from: from ?? { column: 0, row: 0, columnOffset: 0, rowOffset: 0 } };
    if (anchorNode.localName === "twoCellAnchor" && to) anchor.to = to;
    if (ext) { anchor.width = Number(attr(ext, "cx")) || 0; anchor.height = Number(attr(ext, "cy")) || 0; }
    if (pos) { anchor.x = Number(attr(pos, "x")) || 0; anchor.y = Number(attr(pos, "y")) || 0; }
    for (let content of anchorNode.children) {
      if (typeof content === "string") continue;
      // Slicers are graphic frames in the Excel 2010 branch; the fallback is a "not supported" text box.
      const slicer = content.localName === "AlternateContent" ? child(child(child(child(child(content, "Choice"), "graphicFrame"), "graphic"), "graphicData"), "slicer") : undefined;
      if (slicer) {
        drawings.push({ kind: "slicer", anchor, name: attr(slicer, "name") ?? "", columnCount: 1, itemHeight: 241_300, items: [], filtered: false, style: {} });
        continue;
      }
      if (content.localName === "AlternateContent") content =child(content, "Fallback")?.children.find((part): part is XmlNode => typeof part !== "string") ?? child(content, "Choice")?.children.find((part): part is XmlNode => typeof part !== "string") ?? content;
      if (content.localName === "pic") {
        const embed = attr(child(child(content, "blipFill"), "blip"), "embed"), picture = image(embed ?? "");
        if (!picture) continue;
        const name = attr(child(child(content, "nvPicPr"), "cNvPr"), "name");
        drawings.push({ kind: "image", anchor, ...picture, ...(name ? { name } : {}) });
      } else if (content.localName === "sp" || content.localName === "grpSp" || content.localName === "cxnSp") {
        // Shapes and groups keep their own EMU frame; the renderer maps it onto the anchor's cells.
        const xfrm = child(child(content, "spPr") ?? child(content, "grpSpPr"), "xfrm"), offset = child(xfrm, "off"), size = child(xfrm, "ext");
        const frame = size
          ? { left: Number(attr(offset, "x")) || 0, top: Number(attr(offset, "y")) || 0, width: Math.max(1, Number(attr(size, "cx")) || 0), height: Math.max(1, Number(attr(size, "cy")) || 0) }
          : { left: 0, top: 0, width: anchor.width || 914_400, height: anchor.height || 914_400 };
        const elements = parseDrawingShapes({ name: "spTree", localName: "spTree", attributes: {}, children: [content] }, { theme: themeRoot, image, fallbackBox: frame });
        const nonVisual = child(child(content, "nvSpPr") ?? child(content, "nvGrpSpPr") ?? child(content, "nvCxnSpPr"), "cNvPr"), name = attr(nonVisual, "descr") ?? attr(nonVisual, "name");
        if (elements.length) drawings.push({ kind: "shapes", anchor, frame, elements, ...(name ? { name } : {}) });
      }
    }
  }
  return drawings;
}

function imageMimeType(path: string): string | undefined {
  const extension = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
  return ({ png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp", svg: "image/svg+xml" } as Record<string, string>)[extension];
}

// ---------------------------------------------------------------------------------------------------------------------
// Package helpers

function readRelationships(zip: ZipArchive, partPath: string): Map<string, Relationship> {
  const directory = partPath.slice(0, partPath.lastIndexOf("/")), file = partPath.slice(partPath.lastIndexOf("/") + 1);
  const relsPath = `${directory}/_rels/${file}.rels`, result = new Map<string, Relationship>();
  if (!zip.has(relsPath)) return result;
  for (const rel of elements(decodeXml(zip, relsPath), "Relationship")) {
    const id = attr(rel, "Id"), target = attr(rel, "Target"), type = attr(rel, "Type") ?? "";
    if (!id || !target) continue;
    if (attr(rel, "TargetMode") === "External") { result.set(id, { target, type, external: true }); continue; }
    const joined = target.startsWith("/") ? target.slice(1) : `${directory}/${target}`, parts: string[] = [];
    let escaped = false;
    for (const part of joined.split("/")) {
      if (part === "..") { if (!parts.length) { escaped = true; break; } parts.pop(); }
      else if (part && part !== ".") parts.push(part);
    }
    if (!escaped) result.set(id, { target: parts.join("/"), type, external: false });
  }
  return result;
}

export function parseCellReference(ref: string): { row: number; column: number } {
  const match = /^\$?([A-Z]{1,3})\$?([1-9][0-9]*)$/i.exec(ref);
  if (!match) return { row: 0, column: 0 };
  let column = 0; for (const char of match[1]!.toUpperCase()) column = column * 26 + char.charCodeAt(0) - 64;
  const row = Number(match[2]);
  return { row: Number.isSafeInteger(row) && row <= 1_048_576 ? row : 0, column: column <= 16_384 ? column : 0 };
}

function parseRange(ref: string): SpreadsheetRange | undefined {
  const [first, second] = ref.replace(/^.*!/, "").split(":");
  const a = parseCellReference(first ?? ""), b = second ? parseCellReference(second) : a;
  if (a.row && a.column && b.row && b.column) return { top: Math.min(a.row, b.row), left: Math.min(a.column, b.column), bottom: Math.max(a.row, b.row), right: Math.max(a.column, b.column) };
  const columns = /^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})$/i.exec(ref);
  if (columns) { const left = parseCellReference(`${columns[1]}1`).column, right = parseCellReference(`${columns[2]}1`).column; return { top: 1, left: Math.min(left, right), bottom: 1_048_576, right: Math.max(left, right) }; }
  const rows = /^\$?(\d+):\$?(\d+)$/.exec(ref);
  if (rows) return { top: Math.min(Number(rows[1]), Number(rows[2])), left: 1, bottom: Math.max(Number(rows[1]), Number(rows[2])), right: 16_384 };
  return undefined;
}

export function columnName(value: number): string { let name = ""; while (value > 0) { const remainder = (value - 1) % 26; name = String.fromCharCode(65 + remainder) + name; value = Math.floor((value - 1) / 26); } return name; }

function todaySerial(): number { const now = new Date(); return dateToSerial(now.getFullYear(), now.getMonth(), now.getDate()); }
