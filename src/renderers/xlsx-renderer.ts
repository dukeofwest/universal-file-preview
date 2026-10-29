import type { FileFormat } from "../detection/detect-format.js";
import type { SpreadsheetBorderEdge, SpreadsheetCell, SpreadsheetConditional, SpreadsheetDrawing, SpreadsheetFill, SpreadsheetFont, SpreadsheetRange, SpreadsheetRun, SpreadsheetSlicer, SpreadsheetSlicerStyle, SpreadsheetSparkline, SpreadsheetStyle, SpreadsheetTable, Workbook, Worksheet } from "../model.js";
import { parseXlsx } from "../parsers/xlsx.js";
import { formatGeneral } from "../parsers/xlsx-number-format.js";
import { renderPptxElement } from "./drawingml-renderer.js";
import { fontStack } from "./fonts.js";
import type { FileRenderer, RendererContext, RendererInstance } from "./types.js";

const EMU_PER_PX = 9525;
const GRIDLINE = "#e0e0e0";
const MAX_ROWS = 1_048_576, MAX_COLUMNS = 16_384;
type Edge = SpreadsheetBorderEdge;
const thinEdges = (color: string) => ({ left: { style: "thin", color }, right: { style: "thin", color }, top: { style: "thin", color }, bottom: { style: "thin", color } });
/** Slicer parts a style leaves out look like Office's default light slicer. */
const slicerDefaults: SpreadsheetSlicerStyle = {
  header: { font: { bold: true }, borders: { bottom: { style: "thin", color: "#bfbfbf" } } },
  selectedItemWithData: { font: { color: "#000000" }, fill: { type: "pattern", pattern: "solid", foreground: "#dae3f3" }, borders: thinEdges("#8faadc") },
  selectedItemWithNoData: { font: { color: "#a6a6a6" }, fill: { type: "pattern", pattern: "solid", foreground: "#edf1f9" }, borders: thinEdges("#b4c7e7") },
  unselectedItemWithData: { font: { color: "#000000" }, fill: { type: "pattern", pattern: "solid", foreground: "#ffffff" }, borders: thinEdges("#bfbfbf") },
  unselectedItemWithNoData: { font: { color: "#bfbfbf" }, fill: { type: "pattern", pattern: "solid", foreground: "#ffffff" }, borders: thinEdges("#e0e0e0") },
};
type Effective = { style: SpreadsheetStyle; font: SpreadsheetFont; fill?: SpreadsheetFill; borders: Partial<Record<"left" | "right" | "top" | "bottom", Edge>>; conditional?: SpreadsheetConditional };

export class XlsxRenderer implements FileRenderer {
  readonly formats: readonly FileFormat[] = ["xlsx"];
  render({ bytes, file }: RendererContext): RendererInstance {
    const locale = typeof navigator !== "undefined" && navigator.language ? navigator.language : "en-US";
    const workbook = parseXlsx(bytes, file.name, { locale }), root = document.createElement("div"); root.className = "ufp-sheet-view";
    const tabBar = document.createElement("div"); tabBar.className = "ufp-sheet-tabs"; tabBar.setAttribute("role", "tablist");
    const viewport = document.createElement("div"); viewport.className = "ufp-grid-viewport";
    root.append(tabBar, viewport);
    let zoom = 1, view: SheetView | undefined;
    const visible = workbook.sheets.map((sheet, index) => ({ sheet, index })).filter(entry => !entry.sheet.hidden);
    const tabs = visible.length ? visible : workbook.sheets.map((sheet, index) => ({ sheet, index }));
    const selectSheet = (index: number) => {
      view?.destroy();
      view = new SheetView(workbook, workbook.sheets[index]!, zoom, locale);
      viewport.replaceChildren(view.element);
      view.attach();
      [...tabBar.children].forEach(tab => { tab.setAttribute("aria-selected", String(Number((tab as HTMLElement).dataset["sheet"]) === index)); });
    };
    for (const { sheet, index } of tabs) {
      const tab = document.createElement("button"); tab.type = "button"; tab.className = "ufp-sheet-tab"; tab.textContent = sheet.name; tab.dataset["sheet"] = String(index);
      if (sheet.tabColor) tab.style.boxShadow = `inset 0 -3px 0 ${sheet.tabColor}`;
      tab.setAttribute("role", "tab"); tab.onclick = () => selectSheet(index); tabBar.append(tab);
    }
    const initial = tabs.find(entry => entry.index === (workbook.activeSheet ?? 0)) ?? tabs[0]!;
    selectSheet(initial.index);
    return {
      element: root,
      setZoom(value) { zoom = Math.max(0.1, Math.min(4, value)); view?.setZoom(zoom); },
      fitToWidth() { zoom = 1; view?.setZoom(1); },
      destroy() { view?.destroy(); tabBar.replaceChildren(); viewport.replaceChildren(); },
    };
  }
}

class SheetView {
  readonly element: HTMLDivElement;
  private readonly canvas = document.createElement("div");
  private readonly top = document.createElement("div");
  private readonly corner = document.createElement("div");
  private readonly left = document.createElement("div");
  private readonly body = document.createElement("div");
  private readonly cells = new Map<number, SpreadsheetCell>();
  private readonly mergeAt = new Map<number, SpreadsheetRange>();
  private readonly covered = new Map<number, SpreadsheetRange>();
  private readonly columnInfo = new Map<number, { width?: number; hidden?: boolean; style?: number }>();
  private readonly effectiveCache = new Map<number, Effective>();
  private readonly imageUrls: string[] = [];
  private readonly sparklines = new Map<number, SpreadsheetSparkline>();
  private readonly mdw: number;
  private readonly rowCount: number;
  private readonly columnCount: number;
  private readonly frozenRows: number;
  private readonly frozenColumns: number;
  private readonly baseZoom: number;
  private colX = new Float64Array(0);
  private rowY = new Float64Array(0);
  private headerWidth = 0;
  private headerHeight = 0;
  private rendered: { r1: number; r2: number; c1: number; c2: number } | undefined;
  private frame = 0;
  private selected?: { row: number; column: number };
  private pendingTopLeft: { row: number; column: number } | undefined;
  private resizeObserver?: ResizeObserver;
  private readonly onScroll = () => { if (!this.frame) this.frame = requestAnimationFrame(() => { this.frame = 0; this.update(false); }); };
  // Text overflow is measured on a canvas; once substitute font faces have loaded, measure again.
  private readonly onFontsLoaded = () => { syntheticBold.clear(); this.update(true); };

  constructor(private readonly workbook: Workbook, private readonly sheet: Worksheet, private zoom: number, private readonly locale: string) {
    this.element = document.createElement("div"); this.element.className = "ufp-grid-scroll"; this.element.setAttribute("role", "grid"); this.element.setAttribute("aria-label", sheet.name);
    this.canvas.className = "ufp-grid-canvas";
    this.top.className = "ufp-grid-pane-top"; this.corner.className = "ufp-grid-pane-corner"; this.left.className = "ufp-grid-pane-left"; this.body.className = "ufp-grid-body";
    this.top.append(this.corner); this.canvas.append(this.top, this.left, this.body); this.element.append(this.canvas);
    for (const cell of sheet.cells) { const position = cellPosition(cell.ref); if (position) this.cells.set(key(position.row, position.column), cell); }
    for (const range of sheet.merges ?? []) {
      this.mergeAt.set(key(range.top, range.left), range);
      for (let row = range.top; row <= Math.min(range.bottom, range.top + 5000); row++) for (let column = range.left; column <= Math.min(range.right, range.left + 500); column++) if (row !== range.top || column !== range.left) this.covered.set(key(row, column), range);
    }
    for (const sparkline of sheet.sparklines ?? []) { const position = cellPosition(sparkline.ref); if (position) this.sparklines.set(key(position.row, position.column), sparkline); }
    for (const column of sheet.columns ?? []) for (let index = column.min; index <= Math.min(column.max, MAX_COLUMNS, column.min + 16_384); index++) this.columnInfo.set(index, column);
    this.mdw = maxDigitWidth(workbook.defaultFont ?? { family: "Calibri", size: 11 });
    this.baseZoom = sheet.zoom ?? 1;
    this.frozenRows = sheet.frozen?.rows ?? 0; this.frozenColumns = sheet.frozen?.columns ?? 0;
    // Show the used area plus a margin of empty grid, like Excel's endless sheet.
    let lastRow = sheet.maxRow, lastColumn = sheet.maxColumn;
    for (const range of sheet.merges ?? []) { lastRow = Math.max(lastRow, range.bottom); lastColumn = Math.max(lastColumn, range.right); }
    for (const drawing of sheet.drawings ?? []) { const end = drawing.anchor.to ?? drawing.anchor.from; lastRow = Math.max(lastRow, end.row + 2); lastColumn = Math.max(lastColumn, end.column + 2); }
    this.rowCount = Math.min(MAX_ROWS, Math.max(lastRow + 40, 60, this.frozenRows + 40));
    this.columnCount = Math.min(MAX_COLUMNS, Math.max(lastColumn + 12, 26, this.frozenColumns + 12));
    this.pendingTopLeft = sheet.topLeftCell ? cellPosition(sheet.topLeftCell) : undefined;
    this.element.addEventListener("scroll", this.onScroll, { passive: true });
    this.body.addEventListener("click", event => this.select(event));
    this.layout();
  }

  attach(): void {
    if (typeof ResizeObserver !== "undefined") { this.resizeObserver = new ResizeObserver(() => this.update(false)); this.resizeObserver.observe(this.element); }
    document.fonts?.addEventListener("loadingdone", this.onFontsLoaded);
    this.update(true);
  }

  /** Excel opens a sheet scrolled to its saved top-left cell; applied once the grid has a layout box. */
  private applyInitialScroll(): void {
    const topLeft = this.pendingTopLeft;
    if (!topLeft || !this.element.clientHeight) return;
    this.pendingTopLeft = undefined;
    this.element.scrollTop = this.rowY[Math.min(topLeft.row, this.rowCount)]! - this.rowY[1]!;
    this.element.scrollLeft = this.colX[Math.min(topLeft.column, this.columnCount)]! - this.colX[1]!;
  }

  setZoom(zoom: number): void {
    const ratioX = this.element.scrollLeft / Math.max(1, this.colX[this.columnCount + 1]!), ratioY = this.element.scrollTop / Math.max(1, this.rowY[this.rowCount + 1]!);
    this.zoom = zoom; this.layout();
    this.element.scrollLeft = ratioX * this.colX[this.columnCount + 1]!; this.element.scrollTop = ratioY * this.rowY[this.rowCount + 1]!;
    this.update(true);
  }

  destroy(): void {
    this.element.removeEventListener("scroll", this.onScroll);
    document.fonts?.removeEventListener("loadingdone", this.onFontsLoaded);
    this.resizeObserver?.disconnect();
    if (this.frame) cancelAnimationFrame(this.frame);
    for (const url of this.imageUrls) URL.revokeObjectURL(url);
    this.canvas.replaceChildren();
  }

  private get scale(): number { return this.zoom * this.baseZoom; }

  /** Pixel geometry: Excel column width (characters) -> pixels via the maximum digit width of the default font. */
  private layout(): void {
    const z = this.scale, showHeaders = this.sheet.showHeaders !== false;
    const defaultWidth = this.defaultColumnPixels();
    const defaultHeight = this.sheet.defaultRowHeight ?? 15;
    this.colX = new Float64Array(this.columnCount + 2);
    this.rowY = new Float64Array(this.rowCount + 2);
    const digits = String(this.rowCount).length;
    this.headerWidth = showHeaders ? Math.round(Math.max(26, digits * 7 + 12) * z) : 0;
    this.headerHeight = showHeaders ? Math.round(20 * z) : 0;
    this.colX[1] = this.headerWidth;
    for (let column = 1; column <= this.columnCount; column++) {
      const info = this.columnInfo.get(column);
      const width = info?.hidden ? 0 : info?.width !== undefined ? columnPixels(info.width, this.mdw) : defaultWidth;
      this.colX[column + 1] = this.colX[column]! + Math.round(width * z);
    }
    this.rowY[1] = this.headerHeight;
    const autoHeights = this.autoRowHeights(defaultHeight);
    for (let row = 1; row <= this.rowCount; row++) {
      const info = this.sheet.rows?.[row];
      const points = info?.hidden ? 0 : info?.height ?? autoHeights.get(row) ?? defaultHeight;
      // Excel truncates row heights to whole pixels (43.15pt -> 57px).
      this.rowY[row + 1] = this.rowY[row]! + Math.round(Math.floor(points * 4 / 3 + 1e-6) * z);
    }
    this.canvas.style.width = `${this.colX[this.columnCount + 1]}px`;
    this.canvas.style.height = `${this.rowY[this.rowCount + 1]}px`;
    this.effectiveCache.clear();
    this.rendered = undefined;
  }

  private defaultColumnPixels(): number {
    const stored = this.sheet.defaultColumnWidth;
    if (stored !== undefined && stored > 0) return columnPixels(stored, this.mdw);
    const base = stored !== undefined && stored < 0 ? -stored : 8;
    return Math.ceil((base * this.mdw + 5) / 8) * 8;
  }

  /** Rows without a stored height grow to fit larger fonts, as Excel auto-fits them. */
  private autoRowHeights(defaultHeight: number): Map<number, number> {
    const heights = new Map<number, number>(), styles = this.workbook.styles ?? [];
    const defaultSize = this.workbook.defaultFont?.size ?? 11;
    for (const cell of this.sheet.cells) {
      const position = cellPosition(cell.ref);
      if (!position || this.sheet.rows?.[position.row]?.height !== undefined || cell.display === undefined || cell.display === "") continue;
      const base = { ...this.workbook.defaultFont, ...styles[cell.style ?? 0]?.font }, fonts = [base, ...(cell.runs ?? []).map(run => ({ ...base, ...run.font }))];
      if (!fonts.some(font => (font.size ?? defaultSize) > defaultSize)) continue;
      const height = Math.max(...fonts.map(font => cellLinePixels(font, 1))) * 0.75;
      if (height > (heights.get(position.row) ?? defaultHeight)) heights.set(position.row, height);
    }
    return heights;
  }

  private columnAt(x: number): number { return search(this.colX, x, this.columnCount); }
  private rowAt(y: number): number { return search(this.rowY, y, this.rowCount); }

  private update(force: boolean): void {
    if (this.pendingTopLeft) { this.applyInitialScroll(); if (!this.pendingTopLeft) force = true; }
    const width = this.element.clientWidth || 1200, height = this.element.clientHeight || 800;
    const frozenWidth = this.colX[this.frozenColumns + 1]! - this.colX[1]!, frozenHeight = this.rowY[this.frozenRows + 1]! - this.rowY[1]!;
    const x0 = this.element.scrollLeft + this.colX[1]! + frozenWidth, y0 = this.element.scrollTop + this.rowY[1]! + frozenHeight;
    const c1 = Math.max(this.frozenColumns + 1, this.columnAt(x0)), c2 = Math.min(this.columnCount, this.columnAt(this.element.scrollLeft + width) + 1);
    const r1 = Math.max(this.frozenRows + 1, this.rowAt(y0)), r2 = Math.min(this.rowCount, this.rowAt(this.element.scrollTop + height) + 1);
    const current = this.rendered;
    if (!force && current && r1 >= current.r1 && r2 <= current.r2 && c1 >= current.c1 && c2 <= current.c2) return;
    const window = { r1: Math.max(this.frozenRows + 1, r1 - 30), r2: Math.min(this.rowCount, r2 + 30), c1: Math.max(this.frozenColumns + 1, c1 - 8), c2: Math.min(this.columnCount, c2 + 8) };
    this.rendered = window;
    this.render(window);
  }

  private render(window: { r1: number; r2: number; c1: number; c2: number }): void {
    const topHeight = this.rowY[this.frozenRows + 1]!, leftWidth = this.colX[this.frozenColumns + 1]!;
    const fullWidth = this.colX[this.columnCount + 1]!, fullHeight = this.rowY[this.rowCount + 1]!;
    Object.assign(this.top.style, { height: `${topHeight}px`, width: `${fullWidth}px` });
    Object.assign(this.corner.style, { height: `${topHeight}px`, width: `${leftWidth}px` });
    Object.assign(this.left.style, { height: `${fullHeight - topHeight}px`, width: `${leftWidth}px` });
    this.body.replaceChildren(); this.left.replaceChildren(); this.corner.replaceChildren();
    [...this.top.children].forEach(node => { if (node !== this.corner) node.remove(); });
    const frozenRows = { r1: 1, r2: this.frozenRows }, frozenColumns = { c1: 1, c2: this.frozenColumns };
    this.renderRegion(this.body, { r1: window.r1, r2: window.r2 }, { c1: window.c1, c2: window.c2 }, 0);
    if (this.frozenRows) this.renderRegion(this.top, frozenRows, { c1: window.c1, c2: window.c2 }, 0);
    if (this.frozenColumns) this.renderRegion(this.left, { r1: window.r1, r2: window.r2 }, frozenColumns, topHeight);
    if (this.frozenRows && this.frozenColumns) this.renderRegion(this.corner, frozenRows, frozenColumns, 0);
    if (this.sheet.showHeaders !== false) this.renderHeaders(window, topHeight);
    this.renderDrawings(window);
    this.renderSelection();
  }

  private renderHeaders(window: { r1: number; r2: number; c1: number; c2: number }, topHeight: number): void {
    const font = `${Math.round(11 * 4 / 3 * this.scale)}px ${fontStack(this.workbook.defaultFont?.family ?? "Calibri")}`;
    const header = (target: HTMLElement, text: string, x: number, y: number, w: number, h: number, extra = "") => {
      if (w <= 0 || h <= 0) return;
      const item = document.createElement("div"); item.className = `ufp-cell-header${extra}`; item.textContent = text;
      item.style.cssText = `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;font:${font};line-height:${h - 1}px`;
      target.append(item);
    };
    for (let column = window.c1; column <= window.c2; column++) header(this.top, columnName(column), this.colX[column]!, 0, this.colX[column + 1]! - this.colX[column]!, this.headerHeight);
    for (let column = 1; column <= this.frozenColumns; column++) header(this.corner, columnName(column), this.colX[column]!, 0, this.colX[column + 1]! - this.colX[column]!, this.headerHeight);
    for (let row = window.r1; row <= window.r2; row++) header(this.left, String(row), 0, this.rowY[row]! - topHeight, this.headerWidth, this.rowY[row + 1]! - this.rowY[row]!, " ufp-row-header");
    for (let row = 1; row <= this.frozenRows; row++) header(this.corner, String(row), 0, this.rowY[row]!, this.headerWidth, this.rowY[row + 1]! - this.rowY[row]!, " ufp-row-header");
    header(this.corner, "", 0, 0, this.headerWidth, this.headerHeight, " ufp-cell-corner");
  }

  /** Renders fills, gridlines, borders and text for a block of cells into a pane (y offset for panes placed below the top pane). */
  private renderRegion(target: HTMLElement, rows: { r1: number; r2: number }, columns: { c1: number; c2: number }, offsetY: number): void {
    const fills = document.createDocumentFragment(), lines = document.createDocumentFragment(), borders = document.createDocumentFragment(), texts = document.createDocumentFragment();
    const x = (column: number) => this.colX[column]!, y = (row: number) => this.rowY[row]! - offsetY;
    const gridlines = this.sheet.showGridLines !== false;
    if (gridlines) {
      const top = y(rows.r1), bottom = y(rows.r2 + 1), left = x(columns.c1), right = x(columns.c2 + 1);
      for (let column = columns.c1; column <= columns.c2; column++) if (x(column + 1) > x(column)) lines.append(box("ufp-gridline", x(column + 1) - 1, top, 1, bottom - top, `background:${GRIDLINE}`));
      for (let row = rows.r1; row <= rows.r2; row++) if (y(row + 1) > y(row)) lines.append(box("ufp-gridline", left, y(row + 1) - 1, right - left, 1, `background:${GRIDLINE}`));
    }
    // Merged areas whose anchor lies outside the block still need painting.
    const merges = new Set<SpreadsheetRange>();
    for (const range of this.sheet.merges ?? []) if (range.bottom >= rows.r1 && range.top <= rows.r2 && range.right >= columns.c1 && range.left <= columns.c2) merges.add(range);
    const skipText = new Set<number>();
    for (const range of merges) {
      const effective = this.effective(range.top, range.left);
      const rect = { left: x(range.left), top: y(range.top), width: x(range.right + 1) - x(range.left), height: y(range.bottom + 1) - y(range.top) };
      fills.append(this.fillElement(effective, rect, true, offsetY));
      const sparkline = this.sparklines.get(key(range.top, range.left));
      if (sparkline) fills.append(this.sparklineElement(sparkline, rect));
      texts.append(...this.textElements(range.top, range.left, effective, rect, range));
      skipText.add(key(range.top, range.left));
    }
    for (let row = rows.r1; row <= rows.r2; row++) {
      const top = y(row), height = y(row + 1) - top;
      if (height <= 0) continue;
      for (let column = columns.c1; column <= columns.c2; column++) {
        const left = x(column), width = x(column + 1) - left;
        if (width <= 0) continue;
        const cellKey = key(row, column), merge = this.covered.get(cellKey) ?? this.mergeAt.get(cellKey);
        const effective = this.effective(row, column);
        if (!merge) {
          if (effective.fill || effective.conditional?.dataBar) fills.append(this.fillElement(effective, { left, top, width, height }, false, offsetY));
          const sparkline = this.sparklines.get(cellKey);
          if (sparkline) fills.append(this.sparklineElement(sparkline, { left, top, width, height }));
          if (!skipText.has(cellKey)) texts.append(...this.textElements(row, column, effective, { left, top, width, height }));
        }
        this.appendBorders(borders, row, column, effective, x, y, merge);
      }
    }
    target.append(lines, fills, borders, texts);
  }

  private fillElement(effective: Effective, rect: { left: number; top: number; width: number; height: number }, merged: boolean, offsetY: number): HTMLElement {
    // Excel paints a fill over the cell and all four surrounding gridline pixels; unfilled merged areas paint white over their inner gridlines only.
    const element = box("ufp-cell-fill", rect.left - 1, rect.top - 1, rect.width + 1, rect.height + 1, "");
    // Pattern brushes are anchored to the sheet origin (the gridline pixel above/left of A1), so neighbouring cells join seamlessly.
    const css = fillCss(effective.fill, rect.left - this.colX[1]!, rect.top + offsetY - this.rowY[1]!);
    if (css) element.style.cssText += css;
    else if (merged) Object.assign(element.style, { background: "#fff", left: `${rect.left}px`, top: `${rect.top}px`, width: `${rect.width - 1}px`, height: `${rect.height - 1}px` });
    const bar = effective.conditional?.dataBar;
    if (bar) {
      const inset = Math.max(1, Math.round(2 * this.scale)), span = rect.width - inset * 2 - 1;
      const barElement = box("ufp-data-bar", 1 + inset + span * bar.start, 1 + inset, Math.max(0, span * (bar.end - bar.start)), rect.height - inset * 2 - 1, "");
      barElement.style.background = bar.gradient ? `linear-gradient(to ${bar.negative ? "left" : "right"}, ${bar.color}, ${mix(bar.color, "#ffffff", 0.9)})` : bar.color;
      if (bar.border) barElement.style.boxShadow = `inset 0 0 0 1px ${bar.border}`;
      element.append(barElement);
    }
    return element;
  }

  private appendBorders(target: DocumentFragment, row: number, column: number, effective: Effective, x: (column: number) => number, y: (row: number) => number, merge?: SpreadsheetRange): void {
    // Each cell draws its top and left edges (resolved against the neighbour); the last row/column of a block also draws bottom/right.
    const top = y(row), left = x(column), right = x(column + 1), bottom = y(row + 1);
    const leftEdge = merge && column !== merge.left ? undefined : heavier(effective.borders.left, column > 1 ? this.effective(row, column - 1).borders.right : undefined);
    const topEdge = merge && row !== merge.top ? undefined : heavier(effective.borders.top, row > 1 ? this.effective(row - 1, column).borders.bottom : undefined);
    const rightEdge = merge && column !== merge.right ? undefined : heavier(effective.borders.right, this.effective(row, column + 1).borders.left);
    const bottomEdge = merge && row !== merge.bottom ? undefined : heavier(effective.borders.bottom, this.effective(row + 1, column).borders.top);
    if (topEdge) target.append(borderLine(topEdge, left - 1, top - 1, right - left + 1, true));
    if (leftEdge) target.append(borderLine(leftEdge, left - 1, top - 1, bottom - top + 1, false));
    if (rightEdge && (this.rendered && column === this.rendered.c2 || column === this.frozenColumns || !this.isVisibleColumn(column + 1))) target.append(borderLine(rightEdge, right - 1, top - 1, bottom - top + 1, false));
    if (bottomEdge && (this.rendered && row === this.rendered.r2 || row === this.frozenRows || !this.isVisibleRow(row + 1))) target.append(borderLine(bottomEdge, left - 1, bottom - 1, right - left + 1, true));
    const diagonal = effective.style.borders?.diagonal;
    if (diagonal && (effective.style.borders?.diagonalDown || effective.style.borders?.diagonalUp)) {
      const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      svg.setAttribute("class", "ufp-cell-diagonal"); svg.style.cssText = `position:absolute;left:${left}px;top:${top}px;width:${right - left}px;height:${bottom - top}px;overflow:visible;pointer-events:none`;
      for (const [enabled, x1, y1, x2, y2] of [[effective.style.borders.diagonalDown, 0, 0, right - left, bottom - top], [effective.style.borders.diagonalUp, 0, bottom - top, right - left, 0]] as const) {
        if (!enabled) continue;
        const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
        line.setAttribute("x1", String(x1)); line.setAttribute("y1", String(y1)); line.setAttribute("x2", String(x2)); line.setAttribute("y2", String(y2));
        line.setAttribute("stroke", diagonal.color); line.setAttribute("stroke-width", String(borderWeight(diagonal.style)));
        svg.append(line);
      }
      target.append(svg);
    }
  }
  private isVisibleColumn(column: number): boolean { return column <= this.columnCount && this.colX[column + 1]! > this.colX[column]!; }
  private isVisibleRow(row: number): boolean { return row <= this.rowCount && this.rowY[row + 1]! > this.rowY[row]!; }

  private textElements(row: number, column: number, effective: Effective, rect: { left: number; top: number; width: number; height: number }, merge?: SpreadsheetRange): HTMLElement[] {
    const cell = this.cells.get(key(row, column)), conditional = effective.conditional, ref = `${columnName(column)}${row}`;
    const result: HTMLElement[] = [];
    if (this.sheet.comments?.[ref] !== undefined) {
      const size = Math.max(4, Math.round(6 * this.scale));
      const marker = box("ufp-cell-comment", rect.left + rect.width - size - 1, rect.top, size, size, `border-top:${size}px solid #d00;border-left:${size}px solid transparent;width:0;height:0`);
      marker.title = this.sheet.comments[ref]!; result.push(marker);
    }
    if (this.sheet.filterButtons?.includes(ref)) {
      const size = Math.min(rect.height - 2, Math.round(16 * this.scale));
      const button = box("ufp-filter-button", rect.left + rect.width - size - 2, rect.top + rect.height - size - 2, size, size, "");
      button.innerHTML = `<svg viewBox="0 0 16 16" width="100%" height="100%" aria-hidden="true"><rect x="0.5" y="0.5" width="15" height="15" fill="#fff" stroke="#a6a6a6"/><path d="M4.5 6.5h7L8 10.5z" fill="#444"/></svg>`;
      result.push(button);
    }
    if (conditional?.icon) result.push(this.iconElement(conditional.icon, rect));
    if (!cell || conditional?.hideValue) return result;
    let text = (conditional?.display ?? cell.display ?? cell.value).replace(/\r\n?/g, "\n");
    if (cell.formula && cell.display === undefined && cell.value === "") text = "Cannot calculate";
    if (text === "" || text === undefined) return result;
    const style = effective.style, alignment = style.alignment ?? {}, font = effective.font;
    const numeric = isNumericCell(cell);
    let horizontal = alignment.horizontal ?? "general";
    if (horizontal === "general") horizontal = numeric ? "right" : cell.type === "b" || cell.type === "e" ? "center" : "left";
    const rotation = alignment.textRotation ?? 0;
    const wrap = !!(alignment.wrapText || horizontal === "justify" || horizontal === "distributed" || alignment.vertical === "justify" || alignment.vertical === "distributed");
    const z = this.scale, padding = Math.max(1, Math.round(2 * z));
    const indent = (alignment.indent ?? 0) * Math.round(indentStep(this.workbook.defaultFont ?? {}, this.mdw) * z);
    const fontPx = cellFontPixels(font.size ?? 11, z);
    const fontCss = `${font.italic ? "italic " : ""}${font.bold ? "bold " : ""}${fontPx}px ${fontStack(font.family ?? "Calibri")}`;
    // GDI draws a missing bold face by emboldening each glyph a pixel wider; browsers keep the regular advances.
    const boldSpacing = font.bold && isSyntheticBold(font.family ?? "Calibri") ? Math.max(1, Math.round(z)) : 0;
    const measure = (value: string) => measureText(value, fontCss) + value.length * boldSpacing;
    const available = rect.width - padding * 2 - 1 - indent;
    // Numbers that do not fit show "####"; General numbers first drop precision, as Excel does.
    if (numeric && !wrap && !rotation && !alignment.shrinkToFit && !conditional?.display) {
      if (measure(text) > available) {
        const value = Number(cell.value);
        const general = !(this.workbook.styles?.[cell.style ?? 0]?.numberFormat) && Number.isFinite(value);
        let fitted: string | undefined;
        if (general) for (let length = 10; length >= 1 && !fitted; length--) { const candidate = formatGeneral(value, this.locale, length); if (measure(candidate) <= available) fitted = candidate; }
        text = fitted ?? "#".repeat(Math.max(1, Math.floor(available / Math.max(1, measure("#")))));
      }
    }
    let box_ = { ...rect };
    // Text overflows into empty neighbours (not for numbers, wrapped, merged, or rotated text).
    if (!numeric && !wrap && !merge && !rotation && horizontal !== "fill" && !alignment.shrinkToFit) {
      const needed = measure(text) + padding * 2 + indent;
      if (needed > rect.width) box_ = this.overflowBox(row, column, rect, needed, horizontal);
    }
    if (horizontal === "centerContinuous") box_ = this.centerAcross(row, column, rect);
    const element = document.createElement("div"); element.className = "ufp-grid-cell"; element.setAttribute("role", "gridcell");
    element.setAttribute("aria-label", `${ref}: ${text}`);
    const color = cell.displayColor ?? font.color ?? "#000000";
    const verticalAlign = alignment.vertical ?? "bottom";
    const justify = verticalAlign === "top" || verticalAlign === "justify" || verticalAlign === "distributed" ? "flex-start" : verticalAlign === "center" ? "center" : "flex-end";
    element.style.cssText = `left:${box_.left}px;top:${box_.top}px;width:${box_.width - 1}px;height:${box_.height - 1}px;font:${fontCss};color:${color};padding:0 ${padding}px;justify-content:${justify};`
      + `text-align:${horizontal === "right" ? "right" : horizontal === "center" || horizontal === "centerContinuous" ? "center" : horizontal === "justify" || horizontal === "distributed" ? "justify" : "left"};`
      + `white-space:${wrap ? "pre-wrap" : "pre"};${wrap ? `line-height:${cellLinePixels(font, z)}px;` : ""}${boldSpacing ? `letter-spacing:${boldSpacing}px;` : ""}${indent ? (horizontal === "right" ? `padding-right:${padding + indent}px;` : `padding-left:${padding + indent}px;`) : ""}`;
    if (horizontal === "distributed") element.style.textAlignLast = "justify";
    if (box_.width > rect.width && this.sheet.showGridLines !== false && !effective.fill) element.style.background = "#fff";
    const decorations = [font.underline ? "underline" : "", font.strike ? "line-through" : ""].filter(Boolean).join(" ");
    const line = document.createElement("div"); line.className = "ufp-cell-text";
    if (decorations) { line.style.textDecorationLine = decorations; if (font.underline?.startsWith("double")) line.style.textDecorationStyle = "double"; }
    if (alignment.shrinkToFit && !wrap) {
      const width = measure(text);
      if (width > available && width > 0) line.style.fontSize = `${Math.max(1, fontPx * available / width)}px`;
    }
    const link = this.sheet.hyperlinks?.[ref];
    const content = link ? document.createElement("a") : line;
    if (link && content instanceof HTMLAnchorElement) { content.href = link; content.target = "_blank"; content.rel = "noopener noreferrer"; content.style.color = "inherit"; content.style.textDecoration = "inherit"; line.append(content); }
    if (cell.runs && !conditional?.display && cell.display === cell.value) for (const run of cell.runs) content.append(this.runElement(run, font, z));
    else if (cell.repeat && !wrap) {
      line.style.display = "flex";
      const before = document.createElement("span"), after = document.createElement("span"), filler = document.createElement("span");
      before.textContent = text.slice(0, cell.repeat.index); after.textContent = text.slice(cell.repeat.index);
      filler.style.cssText = "flex:1;overflow:hidden;white-space:pre"; filler.textContent = cell.repeat.char === " " ? "" : cell.repeat.char.repeat(200);
      content.append(before, filler, after); if (content !== line) line.append(content);
    } else if (horizontal === "fill") {
      const unit = measure(text), count = unit > 0 ? Math.max(1, Math.floor(available / unit)) : 1;
      content.append(text.repeat(count));
    } else content.append(wrap && text.endsWith("\n") ? `${text}​` : text); // Excel keeps a trailing line break as an empty last line
    if (rotation) this.rotate(element, line, rotation, horizontal, verticalAlign, text, fontPx, box_);
    element.append(line);
    result.push(element);
    return result;
  }

  private runElement(run: SpreadsheetRun, base: SpreadsheetFont, z: number): HTMLElement {
    const span = document.createElement("span"), font = { ...base, ...run.font };
    span.textContent = run.text;
    if (run.font) {
      span.style.font = `${font.italic ? "italic " : ""}${font.bold ? "bold " : ""}${cellFontPixels(font.size ?? 11, z)}px ${fontStack(font.family ?? "Calibri")}`;
      if (font.color) span.style.color = font.color;
      const decorations = [font.underline ? "underline" : "", font.strike ? "line-through" : ""].filter(Boolean).join(" ");
      if (decorations) span.style.textDecoration = decorations;
      if (font.vertAlign) { span.style.verticalAlign = font.vertAlign === "superscript" ? "super" : "sub"; span.style.fontSize = `${Math.round(cellFontPixels(font.size ?? 11, z) * 0.7)}px`; }
    }
    return span;
  }

  /** Rotated text: for 90/-90 degrees the vertical alignment picks the side and the horizontal alignment the end, as in Excel. */
  private rotate(element: HTMLElement, line: HTMLElement, rotation: number, horizontal: string, vertical: string, text: string, fontPx: number, rect: { height: number }): void {
    element.style.whiteSpace = "pre";
    const side = (start: boolean, middle: boolean) => middle ? "center" : start ? "flex-start" : "flex-end";
    if (rotation === 255) {
      line.style.writingMode = "vertical-rl"; line.style.textOrientation = "upright";
      element.style.alignItems = "center";
      if (text.length * fontPx * 1.2 > rect.height) element.style.justifyContent = "flex-start";
      return;
    }
    if (rotation === 90 || rotation === 180) {
      line.style.writingMode = "vertical-rl";
      element.style.flexDirection = "row";
      if (rotation === 90) {
        line.style.transform = "rotate(180deg)";
        element.style.justifyContent = side(vertical === "top", vertical === "center");
        element.style.alignItems = side(horizontal === "right", horizontal === "center" || horizontal === "centerContinuous");
      } else {
        element.style.justifyContent = side(vertical !== "top" && vertical !== "center", vertical === "center");
        element.style.alignItems = side(horizontal !== "right", horizontal === "center" || horizontal === "centerContinuous");
      }
      return;
    }
    // Other angles rotate around the bottom-left (counter-clockwise) or top-left (clockwise) corner; the rotated box is kept inside the cell.
    const degrees = rotation <= 90 ? -rotation : rotation - 90, lineHeight = Math.round(fontPx * 1.22);
    const shift = Math.round(lineHeight * Math.sin(Math.abs(degrees) * Math.PI / 180)) + 2;
    Object.assign(line.style, { position: "absolute", whiteSpace: "pre", lineHeight: `${lineHeight}px`, left: `${shift}px`, transform: `rotate(${degrees}deg)` });
    if (degrees < 0) Object.assign(line.style, { bottom: "1px", transformOrigin: "0 100%" });
    else Object.assign(line.style, { top: "1px", transformOrigin: "0 0" });
  }


  private overflowBox(row: number, column: number, rect: { left: number; top: number; width: number; height: number }, needed: number, horizontal: string): { left: number; top: number; width: number; height: number } {
    const empty = (target: number) => target >= 1 && target <= this.columnCount && !this.hasContent(row, target) && !this.covered.has(key(row, target)) && !this.mergeAt.has(key(row, target));
    let leftColumn = column, rightColumn = column;
    const width = () => this.colX[rightColumn + 1]! - this.colX[leftColumn]!;
    for (let guard = 0; guard < 64 && width() < needed; guard++) {
      if (horizontal === "right") { if (!empty(leftColumn - 1) || leftColumn - 1 <= this.frozenColumns && column > this.frozenColumns) break; leftColumn--; }
      else if (horizontal === "center") {
        const canLeft = empty(leftColumn - 1), canRight = empty(rightColumn + 1);
        if (!canLeft || !canRight) break;
        leftColumn--; rightColumn++;
      } else { if (!empty(rightColumn + 1)) break; rightColumn++; }
    }
    return { left: this.colX[leftColumn]!, top: rect.top, width: this.colX[rightColumn + 1]! - this.colX[leftColumn]!, height: rect.height };
  }

  private centerAcross(row: number, column: number, rect: { left: number; top: number; width: number; height: number }): { left: number; top: number; width: number; height: number } {
    let right = column;
    while (right < this.columnCount && !this.hasContent(row, right + 1) && this.effective(row, right + 1).style.alignment?.horizontal === "centerContinuous") right++;
    return { ...rect, width: this.colX[right + 1]! - this.colX[column]! };
  }

  private hasContent(row: number, column: number): boolean {
    const cell = this.cells.get(key(row, column));
    // Any value blocks text overflow, including formulas that return an empty string.
    return !!cell && ((cell.display ?? cell.value) !== "" || cell.formula !== undefined) || !!this.sheet.conditional?.[`${columnName(column)}${row}`]?.icon;
  }

  private iconElement(icon: { set: string; index: number }, rect: { left: number; top: number; width: number; height: number }): HTMLElement {
    const size = Math.round(Math.min(16 * this.scale, rect.height - 2));
    const element = box("ufp-cell-icon", rect.left + Math.round(2 * this.scale), rect.top + rect.height - size - Math.round(2 * this.scale), size, size, "");
    element.innerHTML = iconSvg(icon.set, icon.index);
    return element;
  }

  private renderDrawings(window: { r1: number; r2: number; c1: number; c2: number }): void {
    void window;
    for (const drawing of this.sheet.drawings ?? []) {
      const rect = this.anchorRect(drawing);
      if (!rect || rect.width < 0 || rect.height < 0 || rect.width + rect.height <= 0) continue;
      if (drawing.kind === "image") {
        const blob = new Blob([drawing.bytes.slice().buffer as ArrayBuffer], { type: drawing.mimeType }), url = URL.createObjectURL(blob);
        this.imageUrls.push(url);
        const image = document.createElement("img"); image.className = "ufp-sheet-image"; image.src = url; image.alt = drawing.name ?? "";
        image.style.cssText = `position:absolute;left:${rect.left}px;top:${rect.top}px;width:${rect.width}px;height:${rect.height}px;z-index:1`;
        this.body.append(image);
        continue;
      }
      if (drawing.kind === "slicer") { this.body.append(this.slicerElement(drawing, rect)); continue; }
      // Elements are laid out in the object's EMU frame; the container maps that frame onto the anchor's cells (lengths in cqw of the container).
      const container = document.createElement("div"); container.className = "ufp-sheet-drawing";
      container.style.cssText = `left:${rect.left}px;top:${rect.top}px;width:${Math.max(1, rect.width)}px;height:${Math.max(1, rect.height)}px`;
      if (drawing.name) { container.setAttribute("role", "img"); container.setAttribute("aria-label", drawing.name); }
      const { frame } = drawing, frameData = { title: "", paragraphs: [], width: frame.width, height: frame.height, elements: drawing.elements };
      for (const element of drawing.elements) container.append(renderPptxElement({ ...element, left: element.left - frame.left, top: element.top - frame.top }, frameData, this.imageUrls, { runFontLines: true }));
      this.body.append(container);
    }
  }

  /** Excel's slicer layout: 1px frame, 29px caption row over a separator, then a grid of buttons (3px gaps) inset 8px; overflowing buttons get a scroll bar. */
  private slicerElement(slicer: SpreadsheetSlicer, rect: { left: number; top: number; width: number; height: number }): HTMLElement {
    const z = this.scale, px = (value: number) => Math.round(value * z), style = slicer.style, defaults = slicerDefaults;
    const font = this.workbook.defaultFont ?? {}, family = fontStack(font.family ?? "Calibri"), fontPx = cellFontPixels(11, z);
    const frame = document.createElement("div"); frame.className = "ufp-sheet-slicer";
    frame.setAttribute("role", "list"); frame.setAttribute("aria-label", slicer.caption ?? slicer.name);
    const border = Math.max(1, px(1)), inset = px(8), gap = Math.max(1, px(3));
    frame.style.cssText = `left:${rect.left}px;top:${rect.top}px;width:${rect.width}px;height:${rect.height}px;border:${border}px solid ${style.wholeTable?.borders?.left?.color ?? "#bfbfbf"};font:${fontPx}px ${family};${fillCss(style.wholeTable?.fill, 0, 0) || "background:#fff;"}`;
    const inner = { width: rect.width - border * 2, height: rect.height - border * 2 };
    let itemsTop = inset;
    if (slicer.caption !== undefined) {
      const header = style.header ?? defaults.header!, height = px(29);
      const caption = box("ufp-slicer-caption", inset, 0, Math.max(0, inner.width - inset * 2 - px(64)), height, `line-height:${height}px;font-weight:${header.font?.bold === false ? "normal" : "bold"};color:${header.font?.color ?? style.wholeTable?.font?.color ?? "#000"}`);
      caption.textContent = slicer.caption;
      const rule = box("ufp-slicer-rule", inset, height, Math.max(0, inner.width - inset * 2), Math.max(1, px(1)), `background:${header.borders?.bottom?.color ?? "#bfbfbf"}`);
      const icons = box("ufp-slicer-icons", inner.width - px(59), Math.round((height - px(16)) / 2), px(48), px(16), "");
      icons.setAttribute("aria-hidden", "true");
      // Multi-select and clear-filter buttons; the clear button's cross turns red while the slicer filters.
      icons.innerHTML = `<svg viewBox="0 0 48 16" width="100%" height="100%"><path d="M1.5 3.5l1.5 1.5 3-3M1.5 8.5l1.5 1.5 3-3M1.5 13.5l1.5 1.5 3-3" fill="none" stroke="#595959"/><path d="M8 4.5h8M8 9.5h8M8 14.5h8" stroke="#595959"/><path d="M32.5 2.5h11l-4.5 5v5l-2-1.2v-3.8z" fill="none" stroke="#7f7f7f"/><path d="M41.5 10.5l4 4M45.5 10.5l-4 4" stroke="${slicer.filtered ? "#e81123" : "#a6a6a6"}" stroke-width="1.2"/></svg>`;
      frame.append(caption, rule, icons);
      itemsTop = height + Math.max(1, px(1)) + px(4);
    }
    const area = { width: inner.width - inset * 2, height: Math.max(0, inner.height - itemsTop - inset) };
    const columns = Math.max(1, slicer.columnCount), itemHeight = Math.max(1, Math.round(Math.round(slicer.itemHeight / EMU_PER_PX) * z));
    const rows = Math.ceil(slicer.items.length / columns), overflow = rows * (itemHeight + gap) - gap > area.height;
    const itemsWidth = overflow ? area.width - px(23) : area.width, itemWidth = (itemsWidth - gap * (columns - 1)) / columns;
    const items = box("ufp-slicer-items", inset, itemsTop, Math.max(0, itemsWidth), area.height, "overflow:hidden");
    slicer.items.forEach((item, index) => {
      const column = index % columns, row = Math.floor(index / columns);
      const left = Math.round(column * (itemWidth + gap)), right = Math.round(column * (itemWidth + gap) + itemWidth), top = row * (itemHeight + gap);
      if (top >= area.height) return;
      const part = item.selected ? (item.hasData ? "selectedItemWithData" : "selectedItemWithNoData") : (item.hasData ? "unselectedItemWithData" : "unselectedItemWithNoData");
      const look = style[part] ?? defaults[part]!, edge = look.borders?.left?.color;
      const button = box("ufp-slicer-item", left, top, right - left, itemHeight, `${fillCss(look.fill, 0, 0)}color:${look.font?.color ?? "#000"};line-height:${itemHeight - (edge ? border * 2 : 0)}px;padding-left:${inset}px;${edge ? `border:${border}px solid ${edge};` : ""}`);
      button.setAttribute("role", "listitem"); button.setAttribute("aria-selected", String(item.selected));
      button.textContent = item.text;
      items.append(button);
    });
    frame.append(items);
    if (overflow) {
      const track = box("ufp-slicer-scroll", inner.width - inset - px(18), itemsTop, px(18), area.height, "background:#f0f0f0");
      const visible = Math.max(0.1, Math.min(1, area.height / (rows * (itemHeight + gap))));
      track.append(box("ufp-slicer-thumb", px(5), px(17), px(8), Math.max(px(8), (area.height - px(34)) * visible), "background:#c2c2c2"));
      frame.append(track);
    }
    return frame;
  }

  /** In-cell line, column or win/loss chart, inset a few pixels from the cell edges. */
  private sparklineElement(sparkline: SpreadsheetSparkline, rect: { left: number; top: number; width: number; height: number }): SVGSVGElement {
    const z = this.scale, width = rect.width - 1, height = rect.height - 1, padX = Math.max(2, Math.round(4 * z)), padY = Math.max(2, Math.round(4 * z));
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("class", "ufp-sparkline"); svg.setAttribute("aria-hidden", "true");
    svg.style.cssText = `position:absolute;left:${rect.left}px;top:${rect.top}px;width:${Math.max(0, width)}px;height:${Math.max(0, height)}px;overflow:hidden;pointer-events:none`;
    const values = sparkline.emptyCells === "zero" ? sparkline.values.map(value => value ?? 0) : sparkline.values;
    const ordered = sparkline.rightToLeft ? [...values].reverse() : values, numbers = ordered.filter((value): value is number => value !== null);
    if (!numbers.length || width <= padX * 2 || height <= padY * 2) return svg;
    const min = sparkline.min ?? Math.min(...numbers), max = sparkline.max ?? Math.max(...numbers), count = ordered.length;
    const y = (value: number) => max === min ? height / 2 : padY + (max - Math.max(min, Math.min(max, value))) / (max - min) * (height - padY * 2);
    const firstIndex = ordered.findIndex(value => value !== null), lastIndex = ordered.length - 1 - [...ordered].reverse().findIndex(value => value !== null);
    const highest = Math.max(...numbers), lowest = Math.min(...numbers);
    const pointColor = (value: number, index: number, fallback: string | undefined) => index === firstIndex && sparkline.firstColor ? sparkline.firstColor : index === lastIndex && sparkline.lastColor ? sparkline.lastColor
      : value === highest && sparkline.highColor ? sparkline.highColor : value === lowest && sparkline.lowColor ? sparkline.lowColor : value < 0 && sparkline.negativeColor ? sparkline.negativeColor : fallback;
    const element = (name: string, attributes: Record<string, string | number>) => { const node = document.createElementNS("http://www.w3.org/2000/svg", name); for (const [attribute, value] of Object.entries(attributes)) node.setAttribute(attribute, String(value)); svg.append(node); return node; };
    if (sparkline.axisColor && min < 0 && max > 0) element("line", { x1: padX, x2: width - padX, y1: Math.round(y(0)) + 0.5, y2: Math.round(y(0)) + 0.5, stroke: sparkline.axisColor, "stroke-width": Math.max(1, z) });
    if (sparkline.type === "line") {
      const x = (index: number) => count === 1 ? width / 2 : padX + index * (width - padX * 2) / (count - 1);
      const strokeWidth = Math.max(1, sparkline.lineWeight * 4 / 3 * z);
      let path = "", pen = false;
      ordered.forEach((value, index) => { if (value === null) { if (sparkline.emptyCells === "gap") pen = false; return; } path += `${pen ? "L" : "M"}${x(index).toFixed(2)} ${y(value).toFixed(2)}`; pen = true; });
      element("path", { d: path, fill: "none", stroke: sparkline.color, "stroke-width": strokeWidth, "stroke-linejoin": "round", "stroke-linecap": "round" });
      ordered.forEach((value, index) => {
        if (value === null) return;
        const color = pointColor(value, index, sparkline.markersColor);
        if (color) element("circle", { cx: x(index), cy: y(value), r: Math.max(1.5, strokeWidth * 1.5), fill: color });
      });
      return svg;
    }
    // Columns grow from zero (or the axis minimum); win/loss bars fill the upper or lower half.
    const slot = (width - padX * 2) / count, barWidth = Math.max(1, slot * 0.8);
    ordered.forEach((value, index) => {
      if (value === null || sparkline.type === "stacked" && value === 0) return;
      const left = padX + index * slot + (slot - barWidth) / 2;
      const [top, bottom] = sparkline.type === "stacked" ? (value > 0 ? [padY, height / 2] : [height / 2, height - padY])
        : [Math.min(y(value), y(Math.max(min, Math.min(max, 0)))), Math.max(y(value), y(Math.max(min, Math.min(max, 0))))];
      element("rect", { x: left, y: top, width: barWidth, height: Math.max(1, bottom - top), fill: pointColor(value, index, undefined) ?? (value < 0 && sparkline.negativeColor ? sparkline.negativeColor : sparkline.color) });
    });
    return svg;
  }

  private anchorRect(drawing: SpreadsheetDrawing): { left: number; top: number; width: number; height: number } | undefined {
    const z = this.scale, anchor = drawing.anchor;
    const point = (marker: { column: number; row: number; columnOffset: number; rowOffset: number }) => ({
      x: this.colX[Math.min(this.columnCount + 1, marker.column + 1)]! + marker.columnOffset / EMU_PER_PX * z,
      y: this.rowY[Math.min(this.rowCount + 1, marker.row + 1)]! + marker.rowOffset / EMU_PER_PX * z,
    });
    if (anchor.x !== undefined && anchor.y !== undefined && !anchor.to) return { left: this.colX[1]! + anchor.x / EMU_PER_PX * z, top: this.rowY[1]! + anchor.y / EMU_PER_PX * z, width: (anchor.width ?? 0) / EMU_PER_PX * z, height: (anchor.height ?? 0) / EMU_PER_PX * z };
    const from = point(anchor.from);
    if (anchor.to) { const to = point(anchor.to); return { left: from.x, top: from.y, width: to.x - from.x, height: to.y - from.y }; }
    return { left: from.x, top: from.y, width: (anchor.width ?? 0) / EMU_PER_PX * z, height: (anchor.height ?? 0) / EMU_PER_PX * z };
  }

  private select(event: MouseEvent): void {
    const bounds = this.canvas.getBoundingClientRect();
    const column = this.columnAt(event.clientX - bounds.left), row = this.rowAt(event.clientY - bounds.top);
    if (!row || !column) return;
    this.selected = { row, column };
    this.renderSelection();
  }

  private renderSelection(): void {
    this.body.querySelector(".ufp-grid-selection")?.remove();
    this.body.querySelector(".ufp-input-message")?.remove();
    if (!this.selected) return;
    const { row, column } = this.selected, range = this.mergeAt.get(key(row, column)) ?? this.covered.get(key(row, column)) ?? { top: row, left: column, bottom: row, right: column };
    const left = this.colX[range.left]!, top = this.rowY[range.top]!;
    const element = box("ufp-grid-selection", left - 1, top - 1, this.colX[range.right + 1]! - left + 1, this.rowY[range.bottom + 1]! - top + 1, "");
    const cell = this.cells.get(key(range.top, range.left));
    element.setAttribute("aria-label", `${columnName(range.left)}${range.top}: ${cell?.display ?? cell?.value ?? ""}`);
    element.setAttribute("aria-selected", "true");
    this.body.append(element);
    const message = this.sheet.inputMessages?.find(entry => entry.ranges.some(candidate => row >= candidate.top && row <= candidate.bottom && column >= candidate.left && column <= candidate.right));
    if (message) {
      // Excel's input message box sits below the selected cell, slightly indented.
      const tip = document.createElement("div"); tip.className = "ufp-input-message"; tip.setAttribute("role", "tooltip");
      tip.style.left = `${left + 8}px`; tip.style.top = `${this.rowY[range.bottom + 1]! + 6}px`;
      if (message.title) { const title = document.createElement("strong"); title.textContent = message.title; tip.append(title, "\n"); }
      tip.append(message.text);
      this.body.append(tip);
    }
  }

  /** Cell format = cell/row/column style, then table style for unformatted properties, then conditional formatting on top. */
  private effective(row: number, column: number): Effective {
    const cacheKey = key(row, column), cached = this.effectiveCache.get(cacheKey);
    if (cached) return cached;
    const styles = this.workbook.styles ?? [];
    const cell = this.cells.get(cacheKey);
    const styleIndex = cell?.style ?? this.sheet.rows?.[row]?.style ?? this.columnInfo.get(column)?.style ?? 0;
    const base = styles[styleIndex] ?? {};
    const defaultFont = this.workbook.defaultFont ?? {};
    let font: SpreadsheetFont = { ...defaultFont, ...base.font };
    let fill = base.fill;
    const borders: Effective["borders"] = { ...pickEdges(base.borders) };
    const table = this.sheet.tables?.find(candidate => row >= candidate.range.top && row <= candidate.range.bottom && column >= candidate.range.left && column <= candidate.range.right);
    if (table) {
      const tableStyle = tableStyleAt(table, row, column);
      if (tableStyle.font && sameFont(base.font, defaultFont)) font = { ...font, ...tableStyle.font };
      else if (tableStyle.font) {
        // Excel treats cell font properties that equal the Normal style's as unformatted, so the table style shows through for them.
        const own = base.font ?? {};
        if (tableStyle.font.color && (own.color ?? "") === (defaultFont.color ?? "")) font = { ...font, color: tableStyle.font.color };
        if (tableStyle.font.bold && !own.bold) font = { ...font, bold: true };
        if (tableStyle.font.italic && !own.italic) font = { ...font, italic: true };
      }
      if (!fill && tableStyle.fill) fill = tableStyle.fill;
      for (const side of ["left", "right", "top", "bottom"] as const) if (!borders[side] && tableStyle.borders?.[side]) borders[side] = tableStyle.borders[side];
    }
    const conditional = this.sheet.conditional?.[`${columnName(column)}${row}`];
    if (conditional?.style) {
      if (conditional.style.font) font = { ...font, ...conditional.style.font };
      if (conditional.style.fill) fill = conditional.style.fill;
      for (const side of ["left", "right", "top", "bottom"] as const) if (conditional.style.borders?.[side]) borders[side] = conditional.style.borders[side];
    }
    const effective: Effective = { style: base, font, borders, ...(fill ? { fill } : {}), ...(conditional ? { conditional } : {}) };
    this.effectiveCache.set(cacheKey, effective);
    return effective;
  }
}

function tableStyleAt(table: SpreadsheetTable, row: number, column: number): SpreadsheetStyle {
  const { range, elements } = table, layers: Array<{ style: SpreadsheetStyle | undefined; region: SpreadsheetRange }> = [];
  const dataTop = range.top + table.headerRows, dataBottom = range.bottom - table.totalsRows;
  const isHeader = row < dataTop, isTotal = row > dataBottom;
  layers.push({ style: elements["wholeTable"], region: range });
  const band = (first: SpreadsheetStyle & { size?: number } | undefined, second: SpreadsheetStyle & { size?: number } | undefined, offset: number, start: number, end: number, rows: boolean) => {
    const size1 = first?.size ?? 1, size2 = second?.size ?? 1, period = size1 + size2, position = offset % period;
    const inFirst = position < size1, bandStart = offset - (inFirst ? position : position - size1), bandEnd = bandStart + (inFirst ? size1 : size2) - 1;
    const from = start + bandStart, to = Math.min(end, start + bandEnd);
    layers.push({ style: inFirst ? first : second, region: rows ? { top: from, bottom: to, left: range.left, right: range.right } : { top: dataTop, bottom: dataBottom, left: from, right: to } });
  };
  if (!isHeader && !isTotal) {
    if (table.showColumnStripes) band(elements["firstColumnStripe"], elements["secondColumnStripe"], column - range.left, range.left, range.right, false);
    if (table.showRowStripes) band(elements["firstRowStripe"], elements["secondRowStripe"], row - dataTop, dataTop, dataBottom, true);
  }
  if (table.showFirstColumn && column === range.left) layers.push({ style: elements["firstColumn"], region: { ...range, left: range.left, right: range.left } });
  if (table.showLastColumn && column === range.right) layers.push({ style: elements["lastColumn"], region: { ...range, left: range.right, right: range.right } });
  if (isHeader) layers.push({ style: elements["headerRow"], region: { ...range, bottom: dataTop - 1 } });
  if (isTotal) layers.push({ style: elements["totalRow"], region: { ...range, top: dataBottom + 1 } });
  if (isHeader && column === range.left) layers.push({ style: elements["firstHeaderCell"], region: { top: range.top, bottom: dataTop - 1, left: column, right: column } });
  if (isHeader && column === range.right) layers.push({ style: elements["lastHeaderCell"], region: { top: range.top, bottom: dataTop - 1, left: column, right: column } });
  if (isTotal && column === range.left) layers.push({ style: elements["firstTotalCell"], region: { top: dataBottom + 1, bottom: range.bottom, left: column, right: column } });
  if (isTotal && column === range.right) layers.push({ style: elements["lastTotalCell"], region: { top: dataBottom + 1, bottom: range.bottom, left: column, right: column } });
  const result: SpreadsheetStyle = {};
  for (const { style, region } of layers) {
    if (!style) continue;
    if (style.font) result.font = { ...result.font, ...style.font };
    if (style.fill) result.fill = style.fill;
    const edges = style.borders as Record<string, Edge | undefined> | undefined;
    if (!edges) continue;
    const borders = { ...result.borders };
    const pick = (outer: string, inner: string, atEdge: boolean) => atEdge ? edges[outer] : edges[inner];
    const leftEdge = pick("left", "insideV", column === region.left), rightEdge = pick("right", "insideV", column === region.right);
    const topEdge = pick("top", "insideH", row === region.top), bottomEdge = pick("bottom", "insideH", row === region.bottom);
    if (leftEdge) borders.left = leftEdge; if (rightEdge) borders.right = rightEdge; if (topEdge) borders.top = topEdge; if (bottomEdge) borders.bottom = bottomEdge;
    result.borders = borders;
  }
  return result;
}

function pickEdges(borders: SpreadsheetStyle["borders"]): Effective["borders"] {
  const result: Effective["borders"] = {};
  if (!borders) return result;
  for (const side of ["left", "right", "top", "bottom"] as const) if (borders[side]) result[side] = borders[side];
  return result;
}

function sameFont(a: SpreadsheetFont | undefined, b: SpreadsheetFont): boolean {
  if (!a) return true;
  return a.family === b.family && a.size === b.size && (a.color ?? "") === (b.color ?? "") && !!a.bold === !!b.bold && !!a.italic === !!b.italic;
}

const weights: Record<string, number> = { hair: 1, thin: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 };
function borderWeight(style: string): number { return weights[style] ?? 1; }
function heavier(a: Edge | undefined, b: Edge | undefined): Edge | undefined {
  if (!a) return b; if (!b) return a;
  return borderWeight(b.style) > borderWeight(a.style) ? b : a;
}

const dashes: Record<string, number[]> = { dotted: [1, 1], hair: [1, 1], dashed: [3, 1], mediumDashed: [9, 3], dashDot: [9, 3, 3, 3], mediumDashDot: [9, 3, 3, 3], dashDotDot: [9, 3, 3, 3, 3, 3], mediumDashDotDot: [9, 3, 3, 3, 3, 3], slantDashDot: [11, 1, 5, 1] };
/** A border segment centred on the gridline pixel; thicker lines grow to both sides. */
function borderLine(edge: Edge, x: number, y: number, length: number, horizontal: boolean): HTMLElement {
  const weight = borderWeight(edge.style), offset = weight === 1 ? 0 : weight === 2 ? -1 : -1;
  const element = horizontal ? box("ufp-cell-border", x, y + offset, length + (weight > 1 ? 1 : 0), weight, "") : box("ufp-cell-border", x + offset, y, weight, length + (weight > 1 ? 1 : 0), "");
  if (edge.style === "double") element.style.cssText += horizontal ? `border-top:1px solid ${edge.color};border-bottom:1px solid ${edge.color}` : `border-left:1px solid ${edge.color};border-right:1px solid ${edge.color}`;
  else if (dashes[edge.style]) {
    const pattern = dashes[edge.style]!, total = pattern.reduce((sum, value) => sum + value, 0);
    let at = 0;
    const stops = pattern.map((value, index) => { const from = at; at += value; return index % 2 === 0 ? `${edge.color} ${from}px ${at}px` : `transparent ${from}px ${at}px`; }).join(",");
    element.style.background = `repeating-linear-gradient(${horizontal ? "to right" : "to bottom"},${stops})`;
    element.style.backgroundSize = horizontal ? `${total}px 100%` : `100% ${total}px`;
  } else element.style.background = edge.color;
  return element;
}

/** 8x8 bitmaps for Excel pattern fills (1 = foreground pixel). */
const patterns: Record<string, (x: number, y: number) => boolean> = {
  mediumGray: (x, y) => (x + y) % 2 === 0,
  darkGray: (x, y) => !(x % 2 === 0 && y % 2 === 0),
  lightGray: (x, y) => y % 2 === 0 && x % 2 === 0,
  gray125: (x, y) => y % 2 === 0 && x % 4 === (y % 4 === 0 ? 0 : 2),
  gray0625: (x, y) => y % 4 === 0 && x % 4 === (y % 8 === 0 ? 0 : 2),
  darkHorizontal: (_x, y) => y % 4 < 2,
  darkVertical: (x) => x % 4 < 2,
  darkDown: (x, y) => ((x - y) % 4 + 4) % 4 < 2,
  darkUp: (x, y) => (x + y) % 4 >= 2,
  darkGrid: (x, y) => x % 4 < 2 === y % 4 < 2,
  darkTrellis: (x, y) => (x + y) % 2 === 0 || y % 2 === 0 && x % 4 === 1,
  lightHorizontal: (_x, y) => y % 4 === 0,
  lightVertical: (x) => x % 4 === 0,
  lightDown: (x, y) => ((x - y) % 4 + 4) % 4 === 0,
  lightUp: (x, y) => (x + y) % 4 === 3,
  lightGrid: (x, y) => x % 4 === 0 || y % 4 === 0,
  lightTrellis: (x, y) => (x + y) % 4 === 0 || ((x - y) % 4 + 4) % 4 === 0,
};
const patternCache = new Map<string, string>();
function fillCss(fill: SpreadsheetFill | undefined, left: number, top: number): string {
  if (!fill) return "";
  if (fill.type === "gradient") {
    const stops = fill.stops.map(stop => `${stop.color} ${Math.round(stop.position * 100)}%`).join(",");
    if (fill.gradient === "path") return `background:radial-gradient(circle at ${(fill.left ?? 0.5) * 100}% ${(fill.top ?? 0.5) * 100}%,${stops});`;
    return `background:linear-gradient(${fill.degree + 90}deg,${stops});`;
  }
  if (fill.pattern === "solid") return `background:${fill.foreground ?? "#000"};`;
  const bitmap = patterns[fill.pattern];
  if (!bitmap) return `background:${fill.background ?? "#fff"};`;
  const cacheKey = `${fill.pattern}|${fill.foreground}|${fill.background}`;
  let url = patternCache.get(cacheKey);
  if (!url) {
    let rects = "";
    for (let y = 0; y < 8; y++) for (let x = 0; x < 8; x++) if (bitmap(x, y)) rects += `<rect x='${x}' y='${y}' width='1' height='1'/>`;
    url = `url("data:image/svg+xml,${encodeURIComponent(`<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8' shape-rendering='crispEdges'><rect width='8' height='8' fill='${fill.background ?? "#ffffff"}'/><g fill='${fill.foreground ?? "#000000"}'>${rects}</g></svg>`)}")`;
    patternCache.set(cacheKey, url);
  }
  return `background:${url};background-size:8px 8px;background-position:${-left}px ${-top}px;`;
}

function mix(a: string, b: string, amount: number): string {
  const parse = (hex: string) => [1, 3, 5].map(index => Number.parseInt(hex.slice(index, index + 2), 16));
  const x = parse(a), y = parse(b);
  return `#${x.map((channel, index) => Math.round(channel + (y[index]! - channel) * amount).toString(16).padStart(2, "0")).join("")}`;
}

/** Conditional-format icon sets drawn as small SVGs (shapes and colors follow Excel's icon families). */
function iconSvg(set: string, index: number): string {
  const count = Number(set[0]) || 3, level = count - 1 - index; // level 0 = best
  const green = "#3e9a6a", yellow = "#f2c14e", red = "#d9534f", gray = "#8c8c8c", black = "#404040";
  const traffic = [green, yellow, red, black];
  const svg = (content: string) => `<svg viewBox="0 0 16 16" width="100%" height="100%" aria-hidden="true">${content}</svg>`;
  const arrow = (angle: number, color: string) => svg(`<g transform="rotate(${angle} 8 8)"><path d="M8 2 L13 8 H10 V14 H6 V8 H3 Z" fill="${color}"/></g>`);
  const arrowAngles = count === 3 ? [0, 90, 180] : count === 4 ? [0, 45, 135, 180] : [0, 45, 90, 135, 180];
  const colorFor = (position: number) => count === 3 ? [green, yellow, red][position]! : count === 4 ? [green, yellow, "#f08c3a", red][position]! : [green, "#9bc84f", yellow, "#f08c3a", red][position]!;
  if (/Arrows/.test(set)) return arrow(arrowAngles[level] ?? 90, /Gray/.test(set) ? gray : colorFor(level));
  if (/Triangles/.test(set)) return svg(level === 0 ? `<path d="M8 3 L14 12 H2 Z" fill="${green}"/>` : level === count - 1 ? `<path d="M8 13 L14 4 H2 Z" fill="${red}"/>` : `<rect x="3" y="7" width="10" height="3" fill="${yellow}"/>`);
  if (/Stars/.test(set)) { const fill = level === 0 ? "#f4b400" : level === 1 ? "url(#h)" : "none"; return svg(`<defs><linearGradient id="h"><stop offset="50%" stop-color="#f4b400"/><stop offset="50%" stop-color="#fff"/></linearGradient></defs><path d="M8 1.5l1.9 4.2 4.6.4-3.5 3 1.1 4.5L8 11.2 3.9 13.6 5 9.1 1.5 6.1l4.6-.4z" fill="${fill}" stroke="#f4b400"/>`); }
  if (/Rating|Boxes/.test(set)) { const filled = count - 1 - level; let bars = ""; for (let bar = 0; bar < 4; bar++) bars += `<rect x="${2 + bar * 3.2}" y="${11 - bar * 3}" width="2.4" height="${3 + bar * 3}" fill="${bar < filled ? "#3f6fb5" : "#d0d7e5"}"/>`; return svg(bars); }
  if (/Quarters/.test(set)) { const portion = (count - 1 - level) / 4; const angle = portion * 2 * Math.PI, x = 8 + 6 * Math.sin(angle), y = 8 - 6 * Math.cos(angle); return svg(`<circle cx="8" cy="8" r="6" fill="#fff" stroke="${black}"/>${portion >= 1 ? `<circle cx="8" cy="8" r="6" fill="${black}"/>` : portion > 0 ? `<path d="M8 8 V2 A6 6 0 ${portion > 0.5 ? 1 : 0} 1 ${x} ${y} Z" fill="${black}"/>` : ""}`); }
  if (/Symbols/.test(set)) {
    const color = [green, yellow, red][level] ?? red;
    const mark = level === 0 ? `<path d="M4.5 8.2l2.3 2.3 4.7-4.8" stroke="#fff" stroke-width="1.8" fill="none"/>` : level === 1 ? `<rect x="7.1" y="3.8" width="1.8" height="5.2" fill="#fff"/><rect x="7.1" y="10.4" width="1.8" height="1.8" fill="#fff"/>` : `<path d="M5.2 5.2l5.6 5.6M10.8 5.2l-5.6 5.6" stroke="#fff" stroke-width="1.8"/>`;
    return svg(set === "3Symbols2" ? mark.replace(/#fff/g, color) : `<circle cx="8" cy="8" r="6.5" fill="${color}"/>${mark}`);
  }
  if (/Flags/.test(set)) return svg(`<path d="M4 2v12" stroke="${black}" stroke-width="1.2"/><path d="M4.6 2.5h8l-2 3 2 3h-8z" fill="${[green, yellow, red][level] ?? red}"/>`);
  if (/Signs/.test(set)) return svg(level === 0 ? `<circle cx="8" cy="8" r="6" fill="${green}"/>` : level === 1 ? `<path d="M8 2 L14 13 H2 Z" fill="${yellow}"/>` : `<path d="M8 2 L14 8 L8 14 L2 8 Z" fill="${red}"/>`);
  const color = traffic[Math.min(level, traffic.length - 1)]!;
  const ring = /TrafficLights2/.test(set) ? `<circle cx="8" cy="8" r="7" fill="${black}"/>` : "";
  return svg(`${ring}<circle cx="8" cy="8" r="${ring ? 5.2 : 6.5}" fill="${count === 4 && level === 3 ? black : color}"/>`);
}

function box(className: string, left: number, top: number, width: number, height: number, css: string): HTMLDivElement {
  const element = document.createElement("div"); element.className = className;
  element.style.cssText = `position:absolute;left:${left}px;top:${top}px;width:${Math.max(0, width)}px;height:${Math.max(0, height)}px;${css}`;
  return element;
}

function isNumericCell(cell: SpreadsheetCell): boolean {
  if (cell.type && cell.type !== "n" && cell.type !== "d") return false;
  return cell.value.trim() !== "" && Number.isFinite(Number(cell.value));
}

function key(row: number, column: number): number { return row * 16_384 + column; }
function cellPosition(ref: string): { row: number; column: number } | undefined {
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/i.exec(ref);
  if (!match) return undefined;
  let column = 0; for (const char of match[1]!.toUpperCase()) column = column * 26 + char.charCodeAt(0) - 64;
  return { row: Number(match[2]), column };
}
function columnName(value: number): string { let name = ""; while (value > 0) { const remainder = (value - 1) % 26; name = String.fromCharCode(65 + remainder) + name; value = Math.floor((value - 1) / 26); } return name; }
/** Largest index i (1-based) with positions[i] <= value. */
function search(positions: Float64Array, value: number, count: number): number {
  let low = 1, high = count;
  while (low < high) { const middle = (low + high + 1) >> 1; if (positions[middle]! <= value) low = middle; else high = middle - 1; }
  return low;
}

/** Excel draws cell text with GDI fonts at whole-pixel sizes (11pt -> 15px at 100%). */
function cellFontPixels(points: number, zoom: number): number { return Math.max(1, Math.round(points * 4 / 3 * zoom)); }

/** Excel's line pitch for wrapped text in a font, which is also its auto-fit row height: the GDI em size times a per-family factor (measured in Excel; Calibri-like otherwise). */
function cellLinePixels(font: SpreadsheetFont, zoom: number): number {
  const factors: Record<string, number> = { arial: 1.31, "times new roman": 1.32, "franklin gothic book": 1.4, "franklin gothic medium": 1.4, "gill sans mt": 1.5 };
  // Fonts whose pitch follows their hinted (VDMX) heights rather than one factor: pixel em size -> pitch, measured in Excel.
  const pitches: Record<string, Record<number, number>> = { "gill sans mt": { 15: 23, 16: 26, 19: 29, 21: 33, 24: 37, 27: 41, 32: 48, 37: 56, 40: 59, 48: 71 } };
  const family = (font.family ?? "Calibri").toLowerCase(), em = Math.round((font.size ?? 11) * 4 / 3), factor = factors[family] ?? 1.34;
  return Math.max(1, Math.round((pitches[family]?.[em] ?? Math.floor(em * factor + 0.01)) * zoom));
}

/** Pixels per indent level at 100%: three hinted spaces of the default font (Calibri 11: 9px, Gill Sans MT 11: 12px). */
function indentStep(font: SpreadsheetFont, mdw: number): number {
  const spaces: Record<string, number> = { "gill sans mt": 0.2778 };
  const space = spaces[(font.family ?? "Calibri").toLowerCase()];
  return space ? 3 * Math.round(Math.round((font.size ?? 11) * 4 / 3) * space) : 9 * mdw / 7;
}

/** Excel's column width conversion: stored width (characters incl. padding) -> pixels at 100%. */
function columnPixels(width: number, mdw: number): number { return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw); }

let measureContext: CanvasRenderingContext2D | null | undefined;
function measureText(text: string, font: string): number {
  if (measureContext === undefined) measureContext = typeof document !== "undefined" ? document.createElement("canvas").getContext("2d") : null;
  if (!measureContext) return text.length * 7;
  measureContext.font = font;
  return measureContext.measureText(text).width;
}

/** True when the browser has no bold face for a family and fakes it (bold and regular advances are equal). */
const syntheticBold = new Map<string, boolean>();
function isSyntheticBold(family: string): boolean {
  let synthetic = syntheticBold.get(family);
  if (synthetic === undefined) {
    const stack = fontStack(family), sample = "Hamburgefonstiv 0123", regular = measureText(sample, `40px ${stack}`);
    synthetic = Boolean(measureContext) && measureText(sample, `bold 40px ${stack}`) === regular;
    syntheticBold.set(family, synthetic);
  }
  return synthetic;
}

/** Maximum digit width of the workbook's default font, which Excel uses as the unit for column widths. */
function maxDigitWidth(font: SpreadsheetFont): number {
  const size = font.size ?? 11, family = font.family ?? "Calibri", px = size * 4 / 3;
  // Digit widths (em) of common Office fonts, for when the font is not installed (e.g. Office cloud fonts such as Franklin Gothic Book).
  const known: Record<string, number> = { calibri: 0.5, "aptos narrow": 0.48, aptos: 0.55, arial: 0.556, "arial narrow": 0.456, corbel: 0.54, cambria: 0.556, "times new roman": 0.5, verdana: 0.636, tahoma: 0.546, "segoe ui": 0.55, consolas: 0.55, "courier new": 0.6, georgia: 0.62, "franklin gothic book": 0.5864, "franklin gothic medium": 0.5864 };
  // Fonts whose hinted digit advances (hdmx) round up at the pixel em size, so Excel's unit differs from the outline width even when installed.
  const hinted: Record<string, number> = { "gill sans mt": 0.5 };
  const hint = hinted[family.toLowerCase()];
  if (hint) return Math.max(1, Math.round(Math.round(px) * hint));
  if (measureText("0", `${px}px "${family}", monospace`) !== measureText("0", `${px}px monospace`) || measureText("0", `${px}px "${family}", serif`) !== measureText("0", `${px}px serif`)) {
    const digits = Math.max(...[..."0123456789"].map(digit => measureText(digit, `${px}px "${family}"`)));
    if (digits > 0) return Math.round(digits);
  }
  return Math.max(1, Math.round(px * (known[family.toLowerCase()] ?? 0.5)));
}
