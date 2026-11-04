import type { DiagnosticCollector } from "../../diagnostics.js";
import type { ExtendedWorksheet } from "../../model-extensions.js";
import { nsAttr, nsChildren, nsFirst, nsText, type NsXmlNode } from "../../xml/ns-tree.js";
import { ODF } from "./namespaces.js";

/**
 * View hints from settings.xml (implementation specification M14). Internal module.
 *
 * These settings are application-specific (LibreOffice's "ooo:view-settings"), not part of the
 * ODF schema, so they are read best-effort: a value that cannot be read is reported once as
 * `odf.settings.ignored` (info) and left at its default. They only change how the grid opens,
 * never what a cell contains.
 */

export type SheetViewHints = {
  frozen?: { rows: number; columns: number };
  /** 0.1 to 4 (10% to 400%); absent at 100%. */
  zoom?: number;
  showGridLines?: boolean;
  showHeaders?: boolean;
};
/** `view`: the view-wide grid and header switches, for sheets without an entry of their own. */
export type ViewHints = { activeTable?: string; view: Pick<SheetViewHints, "showGridLines" | "showHeaders">; tables: Map<string, SheetViewHints> };

const MAX_ROW = 1_048_576, MAX_COLUMN = 16_384;

/** A config:config-item-* child of `node` with the given config:name. */
function named(node: NsXmlNode | undefined, local: string, name: string): NsXmlNode | undefined {
  return node ? nsChildren(node, ODF.config, local).find(child => nsAttr(child, ODF.config, "name") === name) : undefined;
}

/** Reads the first view of an office settings element (from settings.xml or a flat document). Never throws on content. */
export function readViewHints(settings: NsXmlNode, collector: DiagnosticCollector): ViewHints {
  const hints: ViewHints = { view: {}, tables: new Map() };
  const viewSettings = named(settings, "config-item-set", "ooo:view-settings");
  const view = nsFirst(named(viewSettings, "config-item-map-indexed", "Views"), ODF.config, "config-item-map-entry");
  if (!view) return hints;
  // Each value that cannot be read is counted once.
  const ignore = () => collector.report("odf.settings.ignored", { part: "settings.xml" });
  const item = (container: NsXmlNode, name: string) => { const node = named(container, "config-item", name); return node ? nsText(node).trim() : undefined; };
  const integer = (container: NsXmlNode, name: string) => {
    const text = item(container, name);
    if (text === undefined) return undefined;
    if (!/^-?\d+$/.test(text)) { ignore(); return undefined; }
    return Number(text);
  };
  const boolean = (container: NsXmlNode, name: string) => {
    const text = item(container, name);
    if (text === undefined) return undefined;
    if (text !== "true" && text !== "false") { ignore(); return undefined; }
    return text === "true";
  };

  const active = item(view, "ActiveTable");
  if (active) hints.activeTable = active;
  // View-wide grid and header switches apply to sheets that do not set their own.
  const viewGrid = boolean(view, "ShowGrid"), viewHeaders = boolean(view, "HasColumnRowHeaders");
  if (viewGrid !== undefined) hints.view.showGridLines = viewGrid;
  if (viewHeaders !== undefined) hints.view.showHeaders = viewHeaders;
  const tables = named(view, "config-item-map-named", "Tables");
  const entries = tables ? nsChildren(tables, ODF.config, "config-item-map-entry") : [];
  for (const entry of entries) {
    const name = nsAttr(entry, ODF.config, "name");
    if (name === undefined || hints.tables.has(name)) continue;
    const sheet: SheetViewHints = {};
    // Split mode 2 is a frozen split; the positions then count the frozen columns and rows.
    const columns = integer(entry, "HorizontalSplitMode") === 2 ? integer(entry, "HorizontalSplitPosition") ?? 0 : 0;
    const rows = integer(entry, "VerticalSplitMode") === 2 ? integer(entry, "VerticalSplitPosition") ?? 0 : 0;
    if (columns < 0 || rows < 0 || columns >= MAX_COLUMN || rows >= MAX_ROW) ignore();
    else if (columns || rows) sheet.frozen = { rows, columns };
    // ZoomType 0 is a fixed percentage; other types fit the page and depend on the window.
    const zoomType = integer(entry, "ZoomType") ?? 0, zoom = integer(entry, "ZoomValue");
    if (zoom !== undefined && zoomType === 0) {
      if (zoom < 10 || zoom > 400) ignore();
      else if (zoom !== 100) sheet.zoom = zoom / 100;
    }
    const grid = boolean(entry, "ShowGrid") ?? viewGrid, headers = boolean(entry, "HasColumnRowHeaders") ?? viewHeaders;
    if (grid !== undefined) sheet.showGridLines = grid;
    if (headers !== undefined) sheet.showHeaders = headers;
    hints.tables.set(name, sheet);
  }
  return hints;
}

/** Applies hints to the parsed sheets; returns the active sheet's index when it is not the first. */
export function applyViewHints(sheets: ExtendedWorksheet[], hints: ViewHints, collector: DiagnosticCollector): number | undefined {
  for (const sheet of sheets) {
    const sheetHints: SheetViewHints = hints.tables.get(sheet.name) ?? hints.view;
    if (sheetHints.frozen) sheet.frozen = sheetHints.frozen;
    if (sheetHints.zoom !== undefined) sheet.zoom = sheetHints.zoom;
    if (sheetHints.showGridLines !== undefined) sheet.showGridLines = sheetHints.showGridLines;
    if (sheetHints.showHeaders !== undefined) sheet.showHeaders = sheetHints.showHeaders;
  }
  if (hints.activeTable === undefined) return undefined;
  const index = sheets.findIndex(sheet => sheet.name === hints.activeTable);
  if (index < 0) { collector.report("odf.settings.ignored", { part: "settings.xml" }); return undefined; }
  return index > 0 ? index : undefined;
}
