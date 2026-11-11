import type { DiagnosticCollector } from "../../diagnostics.js";
import type { Paragraph, TextRun } from "../../model.js";
import { nsAttr, nsChildren, nsFirst, type NsXmlNode } from "../../xml/ns-tree.js";
import { ODF } from "./namespaces.js";
import { lengthToPoints, odfColor, parseBorder } from "./units.js";

/**
 * Paragraph and text style resolution for OpenDocument text documents (ODF 1.3 Part 3 §16).
 *
 * Internal module: not exported from the package entry point (decision D2). It is a sibling of
 * styles.ts rather than an extension of it: that resolver stores every family but exposes only
 * spreadsheet accessors, and widening it would change code the spreadsheet parsers depend on. The
 * two share units.ts, where the length, colour and border grammars live.
 *
 * Styles come from two places and both are read: the common styles and defaults in styles.xml, and
 * the automatic styles in content.xml that carry a document's direct formatting. A style inherits
 * through style:parent-style-name, and the chain is followed outwards with the nearest declaration
 * winning. Names come from the document, so every lookup is by own property only.
 */

const { style: STYLE, fo: FO, office: OFFICE } = ODF;
/** Longest inheritance chain followed before the document is treated as having a cycle. */
const MAX_PARENTS = 64;

export interface OdtStyles {
  /** Paragraph formatting for a style name, with its inherited properties already applied. */
  paragraph(name: string | undefined): Partial<Paragraph> | undefined;
  /** Character formatting for a style name (the `text` family, or a paragraph style's own text properties). */
  run(name: string | undefined): Partial<TextRun> | undefined;
  /** The default paragraph and character formatting the document declares. */
  readonly defaults: { paragraph?: Partial<Paragraph>; run?: Partial<TextRun> };
  /** The outline level a paragraph style implies, for headings that do not state one. */
  outlineLevel(name: string | undefined): number | undefined;
  /** The levels of a text:list-style, by the name a text:list refers to. */
  list(name: string | undefined): readonly OdtListLevel[] | undefined;
  /** The list style a paragraph style attaches, for lists that name none themselves. */
  listOfParagraph(name: string | undefined): readonly OdtListLevel[] | undefined;
  table(name: string | undefined): OdtTableStyle | undefined;
  cell(name: string | undefined): OdtCellStyle | undefined;
  columnWidth(name: string | undefined): number | undefined;
}

type RawStyle = {
  readonly family: string;
  readonly parent?: string;
  readonly outlineLevel?: number;
  readonly listStyleName?: string;
  readonly paragraphProperties?: NsXmlNode;
  readonly textProperties?: NsXmlNode;
  readonly tableProperties?: NsXmlNode;
  readonly cellProperties?: NsXmlNode;
  readonly columnProperties?: NsXmlNode;
};

const key = (family: string, name: string) => `${family}\u0000${name}`;

export function createOdtStyles(stylesRoot: NsXmlNode | undefined, automaticStyles: NsXmlNode | undefined, collector: DiagnosticCollector): OdtStyles {
  const styles = new Map<string, RawStyle>();
  const defaults = new Map<string, RawStyle>();
  const lists = new Map<string, readonly OdtListLevel[]>();
  const read = (node: NsXmlNode) => {
    for (const entry of nsChildren(node, STYLE, "style")) {
      const name = nsAttr(entry, STYLE, "name"), family = nsAttr(entry, STYLE, "family");
      if (name === undefined || family === undefined) continue;
      if (!styles.has(key(family, name))) styles.set(key(family, name), rawStyle(entry, family));
    }
    for (const entry of nsChildren(node, STYLE, "default-style")) {
      const family = nsAttr(entry, STYLE, "family");
      if (family !== undefined && !defaults.has(family)) defaults.set(family, rawStyle(entry, family));
    }
    for (const entry of nsChildren(node, TEXT, "list-style")) {
      const name = nsAttr(entry, STYLE, "name");
      if (name !== undefined && !lists.has(name)) lists.set(name, readListStyle(entry));
    }
  };
  // content.xml's automatic styles are read first: a document's direct formatting wins over a common
  // style of the same name, which is how producers override what styles.xml declares.
  if (automaticStyles) read(automaticStyles);
  if (stylesRoot) {
    const office = nsFirst(stylesRoot, OFFICE, "styles");
    if (office) read(office);
    const automatic = nsFirst(stylesRoot, OFFICE, "automatic-styles");
    if (automatic) read(automatic);
  }

  const paragraphCache = new Map<string, Partial<Paragraph> | undefined>();
  const runCache = new Map<string, Partial<TextRun> | undefined>();
  let cycles = 0;

  /** The chain for `name`, nearest first, stopping at a cycle or the depth cap. */
  function chain(family: string, name: string): RawStyle[] {
    const out: RawStyle[] = [], seen = new Set<string>();
    let current: string | undefined = name;
    while (current !== undefined && out.length < MAX_PARENTS) {
      const id = key(family, current);
      if (seen.has(id)) { cycles++; collector.report("odf.style.cycle"); break; }
      seen.add(id);
      const style = styles.get(id);
      if (!style) break;
      out.push(style);
      current = style.parent;
    }
    const fallback = defaults.get(family);
    if (fallback) out.push(fallback);
    return out;
  }

  function resolveParagraph(name: string): Partial<Paragraph> | undefined {
    const resolved: Partial<Paragraph> = {};
    // Outwards: the nearest declaration must win, so the farthest is applied first.
    for (const style of chain("paragraph", name).reverse()) Object.assign(resolved, paragraphOf(style.paragraphProperties));
    return Object.keys(resolved).length ? resolved : undefined;
  }

  function resolveRun(family: string, name: string): Partial<TextRun> | undefined {
    const resolved: Partial<TextRun> = {};
    for (const style of chain(family, name).reverse()) Object.assign(resolved, runOf(style.textProperties));
    return Object.keys(resolved).length ? resolved : undefined;
  }

  return {
    paragraph(name) {
      if (name === undefined) return undefined;
      if (paragraphCache.has(name)) return paragraphCache.get(name);
      const resolved = resolveParagraph(name);
      paragraphCache.set(name, resolved);
      return resolved;
    },
    run(name) {
      if (name === undefined) return undefined;
      if (runCache.has(name)) return runCache.get(name);
      // A name may belong to either family: text:style-name on a span is a text style, on a paragraph
      // it is a paragraph style whose own text properties apply to its runs. The family is chosen by
      // which of them declares the name, not by whether the lookup produced anything: a document that
      // declares a style:default-style for the text family resolves every name in that family, which
      // would hide each paragraph style's own text properties. An unknown name stays with "text", so
      // it still picks up that default.
      const family = styles.has(key("text", name)) || !styles.has(key("paragraph", name)) ? "text" : "paragraph";
      const resolved = resolveRun(family, name);
      runCache.set(name, resolved);
      return resolved;
    },
    defaults: defaultFormatting(defaults.get("paragraph")),
    outlineLevel(name) {
      if (name === undefined) return undefined;
      for (const style of chain("paragraph", name)) if (style.outlineLevel !== undefined) return style.outlineLevel;
      return undefined;
    },
    // A Map is used throughout: names come from the document and must never resolve an inherited property.
    list(name) { return name === undefined ? undefined : lists.get(name); },
    listOfParagraph(name) {
      if (name === undefined) return undefined;
      for (const style of chain("paragraph", name)) if (style.listStyleName !== undefined) return lists.get(style.listStyleName);
      return undefined;
    },
    table(name) {
      if (name === undefined) return undefined;
      for (const style of chain("table", name)) { const resolved = tableStyleOf(style.tableProperties); if (resolved) return resolved; }
      return undefined;
    },
    cell(name) {
      if (name === undefined) return undefined;
      const resolved: Record<string, unknown> = {};
      for (const style of chain("table-cell", name).reverse()) Object.assign(resolved, cellStyleOf(style.cellProperties));
      return Object.keys(resolved).length ? (resolved as OdtCellStyle) : undefined;
    },
    columnWidth(name) {
      if (name === undefined) return undefined;
      for (const style of chain("table-column", name)) { const width = columnWidthOf(style.columnProperties); if (width !== undefined) return width; }
      return undefined;
    },
  };

  function rawStyle(entry: NsXmlNode, family: string): RawStyle {
    const parent = nsAttr(entry, STYLE, "parent-style-name");
    const level = Number(nsAttr(entry, STYLE, "default-outline-level"));
    const paragraphProperties = nsFirst(entry, STYLE, "paragraph-properties");
    const textProperties = nsFirst(entry, STYLE, "text-properties");
    const tableProperties = nsFirst(entry, STYLE, "table-properties");
    const cellProperties = nsFirst(entry, STYLE, "table-cell-properties");
    const columnProperties = nsFirst(entry, STYLE, "table-column-properties");
    const listStyleName = nsAttr(entry, STYLE, "list-style-name");
    return {
      family,
      ...(parent !== undefined ? { parent } : {}),
      ...(Number.isInteger(level) && level >= 1 && level <= 10 ? { outlineLevel: level } : {}),
      ...(listStyleName !== undefined ? { listStyleName } : {}),
      ...(paragraphProperties ? { paragraphProperties } : {}),
      ...(textProperties ? { textProperties } : {}),
      ...(tableProperties ? { tableProperties } : {}),
      ...(cellProperties ? { cellProperties } : {}),
      ...(columnProperties ? { columnProperties } : {}),
    };
  }
}

/** Character formatting from a style:text-properties element. */
function runOf(properties: NsXmlNode | undefined): Partial<TextRun> | undefined {
  if (!properties) return undefined;
  const run: Partial<TextRun> = {};
  const weight = nsAttr(properties, FO, "font-weight");
  if (weight !== undefined) run.bold = weight !== "normal" && weight !== "100" && weight !== "200" && weight !== "300";
  const posture = nsAttr(properties, FO, "font-style");
  if (posture !== undefined) run.italic = posture === "italic" || posture === "oblique";
  const underline = nsAttr(properties, STYLE, "text-underline-style");
  if (underline !== undefined && underline !== "none") {
    run.underline = true;
    const style = underlineStyles[underline];
    if (style) run.underlineStyle = style;
    const color = odfColor(nsAttr(properties, STYLE, "text-underline-color"));
    if (color) run.underlineColor = color;
  } else if (underline === "none") run.underline = false;
  const lineThrough = nsAttr(properties, STYLE, "text-line-through-style");
  if (lineThrough !== undefined) run.strike = lineThrough !== "none";
  const size = lengthToPoints(nsAttr(properties, FO, "font-size"));
  if (size !== undefined && size > 0) run.fontSize = size;
  const family = nsAttr(properties, STYLE, "font-name") ?? nsAttr(properties, FO, "font-family");
  if (family) run.fontFamily = family;
  const color = odfColor(nsAttr(properties, FO, "color"));
  if (color) run.color = color;
  const highlight = odfColor(nsAttr(properties, FO, "background-color"));
  if (highlight) run.highlight = highlight;
  const position = nsAttr(properties, STYLE, "text-position");
  if (position !== undefined) {
    const vertical = verticalAlignOf(position);
    if (vertical) run.verticalAlign = vertical;
  }
  const variant = nsAttr(properties, FO, "font-variant");
  if (variant !== undefined) run.smallCaps = variant === "small-caps";
  const transform = nsAttr(properties, FO, "text-transform");
  if (transform !== undefined) run.caps = transform === "uppercase";
  const spacing = lengthToPoints(nsAttr(properties, FO, "letter-spacing"));
  if (spacing !== undefined) run.letterSpacing = spacing;
  return Object.keys(run).length ? run : undefined;
}

const underlineStyles: Readonly<Record<string, NonNullable<TextRun["underlineStyle"]>>> = {
  solid: "solid", dotted: "dotted", dash: "dashed", "long-dash": "dashed", "dot-dash": "dashed", "dot-dot-dash": "dashed", wave: "wavy",
};

/** "super", "sub", or a percentage: anything above the baseline is superscript, below is subscript. */
function verticalAlignOf(value: string): NonNullable<TextRun["verticalAlign"]> | undefined {
  const first = value.trim().split(/\s+/)[0] ?? "";
  if (first === "super") return "superscript";
  if (first === "sub") return "subscript";
  const percent = Number.parseFloat(first);
  if (!Number.isFinite(percent) || percent === 0) return undefined;
  return percent > 0 ? "superscript" : "subscript";
}

/** Paragraph formatting from a style:paragraph-properties element. */
function paragraphOf(properties: NsXmlNode | undefined): Partial<Paragraph> | undefined {
  if (!properties) return undefined;
  const paragraph: Partial<Paragraph> = {};
  const alignment = alignments[nsAttr(properties, FO, "text-align") ?? ""];
  if (alignment) paragraph.alignment = alignment;
  const left = lengthToPoints(nsAttr(properties, FO, "margin-left"));
  if (left !== undefined) paragraph.indentLeft = left;
  const right = lengthToPoints(nsAttr(properties, FO, "margin-right"));
  if (right !== undefined) paragraph.indentRight = right;
  const firstLine = lengthToPoints(nsAttr(properties, FO, "text-indent"));
  if (firstLine !== undefined) paragraph.indentFirstLine = firstLine;
  const before = lengthToPoints(nsAttr(properties, FO, "margin-top"));
  if (before !== undefined) paragraph.spacingBefore = before;
  const after = lengthToPoints(nsAttr(properties, FO, "margin-bottom"));
  if (after !== undefined) paragraph.spacingAfter = after;
  const background = odfColor(nsAttr(properties, FO, "background-color"));
  if (background) paragraph.background = background;
  if (nsAttr(properties, FO, "break-before") === "page") paragraph.pageBreakBefore = true;
  const lineHeight = nsAttr(properties, FO, "line-height");
  if (lineHeight !== undefined && lineHeight !== "normal") {
    const percent = /%$/.test(lineHeight) ? Number.parseFloat(lineHeight) : undefined;
    if (percent !== undefined && Number.isFinite(percent) && percent > 0) { paragraph.lineHeight = percent / 100; paragraph.lineHeightRule = "auto"; }
    else {
      const exact = lengthToPoints(lineHeight);
      if (exact !== undefined && exact > 0) { paragraph.lineHeight = exact; paragraph.lineHeightRule = "exact"; }
    }
  }
  return Object.keys(paragraph).length ? paragraph : undefined;
}

const alignments: Readonly<Record<string, NonNullable<Paragraph["alignment"]>>> = {
  start: "left", left: "left", end: "right", right: "right", center: "center", justify: "justify",
};

/** The document default paragraph style, as the fields the model actually carries. */
function defaultFormatting(style: RawStyle | undefined): OdtStyles["defaults"] {
  const paragraph = paragraphOf(style?.paragraphProperties), run = runOf(style?.textProperties);
  return { ...(paragraph ? { paragraph } : {}), ...(run ? { run } : {}) };
}

// ---------------------------------------------------------------------------------------------
// List, table and cell styles (milestone 3.3). Read from the same two places as paragraph and
// text styles, and resolved on the same rules.
// ---------------------------------------------------------------------------------------------

/** One level of a list style: how its label is built and how far its text is indented. */
export interface OdtListLevel {
  readonly kind: "ordered" | "bullet";
  /** style:num-format, for example "1", "A", "a", "I" or "i". Absent for a bullet. */
  readonly numFormat?: string;
  readonly prefix?: string;
  readonly suffix?: string;
  readonly startValue: number;
  /** How many enclosing levels the label repeats, so 2 gives "1.2". */
  readonly displayLevels: number;
  readonly bullet?: string;
  readonly indentLeft?: number;
  readonly hanging?: number;
}

export interface OdtTableStyle {
  readonly width?: number;
  readonly widthPercent?: number;
  readonly alignment?: "left" | "center" | "right";
  readonly background?: string;
}
export interface OdtCellStyle {
  readonly background?: string;
  readonly verticalAlign?: "top" | "center" | "bottom";
  readonly borders?: Partial<Record<"top" | "left" | "bottom" | "right", { color: string; width: number; style: string }>>;
  readonly margins?: { top?: number; right?: number; bottom?: number; left?: number };
}

const { text: TEXT, table: TABLE } = ODF;
/** Longest label one list level may produce, so a document cannot build an unbounded string. */
const MAX_LABEL_LENGTH = 64;

/** The levels of a text:list-style element, indexed by level number, which starts at 1. */
export function readListStyle(node: NsXmlNode): readonly OdtListLevel[] {
  const levels: OdtListLevel[] = [];
  for (const child of node.children) {
    if (typeof child === "string" || child.uri !== TEXT) continue;
    const ordered = child.local === "list-level-style-number";
    if (!ordered && child.local !== "list-level-style-bullet" && child.local !== "list-level-style-image") continue;
    const level = Number(nsAttr(child, TEXT, "level") ?? "1");
    if (!Number.isInteger(level) || level < 1 || level > 10) continue;
    const start = Number(nsAttr(child, TEXT, "start-value") ?? "1");
    const display = Number(nsAttr(child, TEXT, "display-levels") ?? "1");
    const properties = nsFirst(child, STYLE, "list-level-properties");
    const alignment = properties ? nsFirst(properties, STYLE, "list-level-label-alignment") : undefined;
    const indentLeft = lengthToPoints(alignment ? nsAttr(alignment, FO, "margin-left") : properties ? nsAttr(properties, TEXT, "space-before") : undefined);
    const hanging = lengthToPoints(alignment ? nsAttr(alignment, FO, "text-indent") : properties ? nsAttr(properties, TEXT, "min-label-width") : undefined);
    const numFormat = nsAttr(child, STYLE, "num-format");
    const prefix = nsAttr(child, STYLE, "num-prefix");
    const suffix = nsAttr(child, STYLE, "num-suffix");
    const bullet = nsAttr(child, TEXT, "bullet-char");
    levels[level] = {
      kind: ordered && numFormat ? "ordered" : "bullet",
      ...(ordered && numFormat ? { numFormat } : {}),
      ...(prefix !== undefined ? { prefix: prefix.slice(0, MAX_LABEL_LENGTH) } : {}),
      ...(suffix !== undefined ? { suffix: suffix.slice(0, MAX_LABEL_LENGTH) } : {}),
      startValue: Number.isInteger(start) ? Math.min(Math.max(start, 0), 1_000_000) : 1,
      displayLevels: Number.isInteger(display) ? Math.min(Math.max(display, 1), 10) : 1,
      ...(bullet ? { bullet: bullet.slice(0, 8) } : {}),
      ...(indentLeft !== undefined ? { indentLeft } : {}),
      ...(hanging !== undefined ? { hanging: Math.abs(hanging) } : {}),
    };
  }
  return levels;
}

/** The label for `level`, given the counters of the enclosing levels: "2.1)", "iv." or a bullet. */
export function listLabel(levels: readonly OdtListLevel[], level: number, counters: readonly number[]): string {
  const definition = levels[level] ?? levels.filter(Boolean)[0];
  if (!definition) return "•";
  if (definition.kind === "bullet" || definition.numFormat === undefined) return definition.bullet ?? "•";
  const from = Math.max(1, level - definition.displayLevels + 1);
  const parts: string[] = [];
  for (let index = from; index <= level; index++) {
    parts.push(formatNumber(counters[index] ?? definition.startValue, levels[index]?.numFormat ?? definition.numFormat));
  }
  return `${definition.prefix ?? ""}${parts.join(".")}${definition.suffix ?? ""}`.slice(0, MAX_LABEL_LENGTH);
}

function formatNumber(value: number, format: string): string {
  const safe = Number.isSafeInteger(value) && value >= 0 ? Math.min(value, 1_000_000) : 0;
  if (format === "A") return alpha(safe).toUpperCase();
  if (format === "a") return alpha(safe);
  if (format === "I") return roman(safe).toUpperCase();
  if (format === "i") return roman(safe);
  return String(safe);
}

/** 1 becomes "a", 26 "z" and 27 "aa", the way spreadsheet column names are built. */
function alpha(value: number): string {
  let out = "", remaining = value;
  while (remaining > 0) {
    const index = (remaining - 1) % 26;
    out = String.fromCharCode(97 + index) + out;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return out || "a";
}

const ROMAN: ReadonlyArray<readonly [number, string]> = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
function roman(value: number): string {
  if (value <= 0 || value > 100_000) return String(value);
  let remaining = value, out = "";
  for (const [amount, numeral] of ROMAN) while (remaining >= amount) { out += numeral; remaining -= amount; }
  return out;
}

/** Table formatting from a style:table-properties element. */
export function tableStyleOf(properties: NsXmlNode | undefined): OdtTableStyle | undefined {
  if (!properties) return undefined;
  const style: Record<string, unknown> = {};
  const points = lengthToPoints(nsAttr(properties, STYLE, "width"));
  if (points !== undefined) style["width"] = points;
  const relative = nsAttr(properties, STYLE, "rel-width");
  if (relative !== undefined) {
    const percent = Number.parseFloat(relative);
    if (Number.isFinite(percent) && percent > 0) style["widthPercent"] = Math.min(percent, 100);
  }
  const align = nsAttr(properties, TABLE, "align");
  if (align === "left" || align === "center" || align === "right") style["alignment"] = align;
  const background = odfColor(nsAttr(properties, FO, "background-color"));
  if (background) style["background"] = background;
  return Object.keys(style).length ? (style as OdtTableStyle) : undefined;
}

/** Cell formatting from a style:table-cell-properties element. */
export function cellStyleOf(properties: NsXmlNode | undefined): OdtCellStyle | undefined {
  if (!properties) return undefined;
  const style: Record<string, unknown> = {};
  const background = odfColor(nsAttr(properties, FO, "background-color"));
  if (background) style["background"] = background;
  const vertical = nsAttr(properties, STYLE, "vertical-align");
  if (vertical === "top" || vertical === "bottom") style["verticalAlign"] = vertical;
  else if (vertical === "middle") style["verticalAlign"] = "center";
  const all = parseBorder(nsAttr(properties, FO, "border"));
  const borders: Record<string, { color: string; width: number; style: string }> = {};
  for (const side of ["top", "left", "bottom", "right"] as const) {
    const edge = parseBorder(nsAttr(properties, FO, `border-${side}`)) ?? all;
    if (edge) borders[side] = { color: edge.color, width: edge.widthPt, style: edge.style };
  }
  if (Object.keys(borders).length) style["borders"] = borders;
  const padding = lengthToPoints(nsAttr(properties, FO, "padding"));
  const margins: Record<string, number> = {};
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const value = lengthToPoints(nsAttr(properties, FO, `padding-${side}`)) ?? padding;
    if (value !== undefined) margins[side] = value;
  }
  if (Object.keys(margins).length) style["margins"] = margins;
  return Object.keys(style).length ? (style as OdtCellStyle) : undefined;
}

/** The column width a style:table-column-properties element declares, in points. */
export function columnWidthOf(properties: NsXmlNode | undefined): number | undefined {
  return properties ? lengthToPoints(nsAttr(properties, STYLE, "column-width")) : undefined;
}
