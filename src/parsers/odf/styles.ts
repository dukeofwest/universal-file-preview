import type { DiagnosticCollector } from "../../diagnostics.js";
import type { SpreadsheetAlignment, SpreadsheetBorderEdge, SpreadsheetBorders, SpreadsheetFont } from "../../model.js";
import type { ExtendedSpreadsheetStyle } from "../../model-extensions.js";
import { nsAttr, nsChildren, type NsXmlNode } from "../../xml/ns-tree.js";
import { EXT, ODF } from "./namespaces.js";
import { createNumberStyles, type OdfNumberFormat, type OdfNumberStyles } from "./number-styles.js";
import type { TextFontSource } from "./text.js";
import { lengthToPoints, odfColor, parseBorder, type OdfBorder } from "./units.js";

/**
 * OpenDocument style resolution for spreadsheets (ODF 1.3 Part 3 §16; implementation
 * specification §4.5, §6.6 and §7.5). Internal module.
 *
 * A style's properties are resolved attribute by attribute: the family's default style, then
 * its parent chain from the root down, then the style itself. Names are scoped by family, and
 * content automatic styles take precedence over common styles of the same name. Cycles in the
 * parent chain are reported and cut.
 */

export interface OdfStyleResolver extends TextFontSource {
  /** Index into `styles` for a table-cell style name; 0 (the default cell style) when unnamed or unknown. */
  cellStyleIndex(name: string | undefined): number;
  /** style:column-width of a table-column style, in points. */
  columnWidthPoints(name: string | undefined): number | undefined;
  /** style:row-height of a table-row style, in points. */
  rowProperties(name: string | undefined): { heightPt?: number };
  /** table:display and tab colour of a table style. */
  tableProperties(name: string | undefined): { hidden: boolean; tabColor?: string };
  /** style:data-style-name of a table-cell style, through its parent chain. */
  dataStyleName(cellStyleName: string | undefined): string | undefined;
  /** The data style of a resolved cell style (an index into `styles`). */
  numberFormat(styleIndex: number): OdfNumberFormat | undefined;
  /** Resolved cell styles; index 0 is the default cell style. Grows as styles are requested. */
  readonly styles: ExtendedSpreadsheetStyle[];
  readonly defaultFont: SpreadsheetFont;
  /** BCP 47 tag from the default cell style's fo:language and fo:country. */
  readonly locale?: string;
}

/** Property attributes of one style, keyed by property element local name, then by "{uri}local". */
type Properties = Map<string, Map<string, string>>;
type RawStyle = { family: string; name: string; parent?: string; dataStyle?: string; properties: Properties };
type Resolved = { properties: Properties; dataStyle?: string };

const MAX_CHAIN = 64;
const key = (uri: string, local: string) => `{${uri}}${local}`;
const familyKey = (family: string, name: string) => `${family}\u0000${name}`;

function readProperties(style: NsXmlNode): Properties {
  const properties: Properties = new Map();
  for (const child of style.children) {
    if (typeof child === "string" || child.uri !== ODF.style || !child.local.endsWith("-properties")) continue;
    const values = properties.get(child.local) ?? new Map<string, string>();
    for (const attribute of child.attributes) values.set(key(attribute.uri, attribute.local), attribute.value);
    properties.set(child.local, values);
  }
  return properties;
}

function readStyle(node: NsXmlNode): RawStyle | undefined {
  const family = nsAttr(node, ODF.style, "family"), name = nsAttr(node, ODF.style, "name");
  if (!family || name === undefined) return undefined;
  const parent = nsAttr(node, ODF.style, "parent-style-name"), dataStyle = nsAttr(node, ODF.style, "data-style-name");
  return { family, name, properties: readProperties(node), ...(parent !== undefined ? { parent } : {}), ...(dataStyle !== undefined ? { dataStyle } : {}) };
}

/**
 * @param stylesXml the office:document-styles root of styles.xml (its office:styles are read)
 * @param automaticStyles office:automatic-styles from content.xml (or a flat document)
 * @param fontFaces every office:font-face-decls element in the document
 */
export function createStyleResolver(stylesXml: NsXmlNode | undefined, automaticStyles: NsXmlNode | undefined, fontFaces: readonly NsXmlNode[], collector: DiagnosticCollector): OdfStyleResolver {
  const common = new Map<string, RawStyle>(), automatic = new Map<string, RawStyle>(), defaults = new Map<string, Properties>();
  const commonStyles = stylesXml ? nsChildren(stylesXml, ODF.office, "styles")[0] : undefined;
  for (const child of commonStyles?.children ?? []) {
    if (typeof child === "string" || child.uri !== ODF.style) continue;
    if (child.local === "default-style") { const family = nsAttr(child, ODF.style, "family"); if (family && !defaults.has(family)) defaults.set(family, readProperties(child)); }
    else if (child.local === "style") { const style = readStyle(child); if (style && !common.has(familyKey(style.family, style.name))) common.set(familyKey(style.family, style.name), style); }
  }
  for (const child of automaticStyles?.children ?? []) {
    if (typeof child === "string" || child.uri !== ODF.style || child.local !== "style") continue;
    const style = readStyle(child);
    if (style && !automatic.has(familyKey(style.family, style.name))) automatic.set(familyKey(style.family, style.name), style);
  }

  const faces = new Map<string, string>();
  for (const declarations of fontFaces) for (const face of nsChildren(declarations, ODF.style, "font-face")) {
    const name = nsAttr(face, ODF.style, "name"), family = nsAttr(face, ODF.svg, "font-family");
    if (name !== undefined && family !== undefined && !faces.has(name)) faces.set(name, unquote(family));
  }

  const resolvedCache = new Map<string, Resolved | null>();
  /** Merges default-style, the parent chain and the style itself; null when the style does not exist. */
  const resolve = (family: string, name: string): Resolved | null => {
    const cacheKey = familyKey(family, name), cached = resolvedCache.get(cacheKey);
    if (cached !== undefined) return cached;
    const chain: RawStyle[] = [], seen = new Set<string>();
    let current = automatic.get(cacheKey) ?? common.get(cacheKey);
    while (current) {
      if (seen.has(current.name) || chain.length >= MAX_CHAIN) { collector.report("odf.style.cycle", { part: "styles" }); break; }
      seen.add(current.name); chain.unshift(current);
      // Parents are common styles (ODF 1.3 Part 3 §16.2); fall back to automatic ones for leniency.
      current = current.parent === undefined ? undefined : common.get(familyKey(family, current.parent)) ?? automatic.get(familyKey(family, current.parent));
    }
    if (!chain.length) { resolvedCache.set(cacheKey, null); return null; }
    const properties = mergeInto(new Map(), defaults.get(family));
    let dataStyle: string | undefined;
    for (const style of chain) { mergeInto(properties, style.properties); if (style.dataStyle !== undefined) dataStyle = style.dataStyle; }
    const result: Resolved = dataStyle !== undefined ? { properties, dataStyle } : { properties };
    resolvedCache.set(cacheKey, result);
    return result;
  };

  // The default cell style: default-style table-cell, then the common style named "Default".
  const defaultResolved = resolve("table-cell", "Default") ?? (defaults.has("table-cell") ? { properties: mergeInto(new Map(), defaults.get("table-cell")) } : null);
  if (!defaultResolved) collector.report("odf.style.default-missing", { part: "styles.xml" });
  const locale = defaultResolved ? localeOf(defaultResolved.properties) : undefined;
  // Data styles: content automatic styles take precedence over common ones, as for cell styles.
  const numberStyles = createNumberStyles([automaticStyles, commonStyles], locale);
  const defaultStyle = defaultResolved ? toCellStyle(defaultResolved, faces, numberStyles) : {};
  const styles: ExtendedSpreadsheetStyle[] = [defaultStyle];
  const defaultFont: SpreadsheetFont = { ...defaultStyle.font, family: defaultStyle.font?.family ?? "Liberation Sans", size: defaultStyle.font?.size ?? 10 };

  const cellIndexes = new Map<string, number>();
  return {
    styles, defaultFont,
    ...(locale !== undefined ? { locale } : {}),
    cellStyleIndex(name) {
      // styles[0] already is the default style merged with the common "Default" style.
      if (name === undefined || (name === "Default" && defaultResolved)) return 0;
      const known = cellIndexes.get(name);
      if (known !== undefined) return known;
      const resolved = resolve("table-cell", name);
      const index = resolved ? styles.push(toCellStyle(resolved, faces, numberStyles)) - 1 : 0;
      cellIndexes.set(name, index);
      return index;
    },
    columnWidthPoints(name) { return name === undefined ? undefined : lengthToPoints(property(resolve("table-column", name), "table-column-properties", ODF.style, "column-width")); },
    rowProperties(name) {
      const height = name === undefined ? undefined : lengthToPoints(property(resolve("table-row", name), "table-row-properties", ODF.style, "row-height"));
      return height !== undefined ? { heightPt: height } : {};
    },
    tableProperties(name) {
      const resolved = name === undefined ? null : resolve("table", name);
      const tabColor = odfColor(property(resolved, "table-properties", ODF.table, "tab-color") ?? property(resolved, "table-properties", EXT.tableooo, "tab-color"));
      return { hidden: property(resolved, "table-properties", ODF.table, "display") === "false", ...(tabColor !== undefined ? { tabColor } : {}) };
    },
    numberFormat(styleIndex) {
      const name = styles[styleIndex]?.numberFormatInfo?.sourceName;
      return name === undefined ? undefined : numberStyles.get(name);
    },
    dataStyleName(cellStyleName) { return cellStyleName === undefined ? undefined : resolve("table-cell", cellStyleName)?.dataStyle; },
    textRunFont(styleName) {
      const resolved = resolve("text", styleName);
      return resolved?.properties.has("text-properties") ? fontOf(resolved.properties, faces) : undefined;
    },
  };
}

function mergeInto(target: Properties, source: Properties | undefined): Properties {
  for (const [element, values] of source ?? []) {
    const merged = target.get(element) ?? new Map<string, string>();
    for (const [name, value] of values) merged.set(name, value);
    target.set(element, merged);
  }
  return target;
}

function property(resolved: Resolved | null | undefined, element: string, uri: string, local: string): string | undefined {
  return resolved?.properties.get(element)?.get(key(uri, local));
}

function unquote(value: string): string { return value.trim().replace(/^(['"])(.*)\1$/, "$2"); }

function fontOf(properties: Properties, faces: ReadonlyMap<string, string>): SpreadsheetFont {
  const text = properties.get("text-properties"), get = (uri: string, local: string) => text?.get(key(uri, local));
  const font: SpreadsheetFont = {};
  const faceName = get(ODF.style, "font-name"), family = (faceName !== undefined ? faces.get(faceName) : undefined) ?? (get(ODF.fo, "font-family") !== undefined ? unquote(get(ODF.fo, "font-family")!) : undefined);
  if (family) font.family = family;
  const size = lengthToPoints(get(ODF.fo, "font-size"));
  if (size !== undefined && size > 0) font.size = size;
  const weight = get(ODF.fo, "font-weight");
  if (weight === "bold" || (weight !== undefined && /^\d+$/.test(weight) && Number(weight) >= 600)) font.bold = true;
  const fontStyle = get(ODF.fo, "font-style");
  if (fontStyle === "italic" || fontStyle === "oblique") font.italic = true;
  const underline = get(ODF.style, "text-underline-style");
  if (underline !== undefined && underline !== "none") font.underline = get(ODF.style, "text-underline-type") === "double" ? "double" : "single";
  const strike = get(ODF.style, "text-line-through-style");
  if (strike !== undefined && strike !== "none") font.strike = true;
  const color = odfColor(get(ODF.fo, "color"));
  if (color) font.color = color;
  const position = get(ODF.style, "text-position")?.trim().split(/\s+/)[0];
  if (position === "super" || (position && /^\d/.test(position) && parseFloat(position) > 0)) font.vertAlign = "superscript";
  else if (position === "sub" || (position && /^-\d/.test(position) && parseFloat(position) < 0)) font.vertAlign = "subscript";
  return font;
}

/** ODF border to a spreadsheet edge (specification §6.6). */
function edge(border: OdfBorder): SpreadsheetBorderEdge {
  const style = border.style === "double" ? "double" : border.style === "dashed" ? "dashed" : border.style === "dotted" ? "dotted" : border.widthPt <= 1 ? "thin" : border.widthPt <= 2 ? "medium" : "thick";
  return { style, color: border.color };
}

function toCellStyle(resolved: Resolved, faces: ReadonlyMap<string, string>, numberStyles: OdfNumberStyles): ExtendedSpreadsheetStyle {
  const style: ExtendedSpreadsheetStyle = {};
  if (resolved.properties.has("text-properties")) {
    const font = fontOf(resolved.properties, faces);
    if (Object.keys(font).length) style.font = font;
  }
  const cell = resolved.properties.get("table-cell-properties"), get = (uri: string, local: string) => cell?.get(key(uri, local));
  const background = odfColor(get(ODF.fo, "background-color"));
  if (background) style.fill = { type: "pattern", pattern: "solid", foreground: background };

  const borders: SpreadsheetBorders = {}, all = parseBorder(get(ODF.fo, "border"));
  for (const side of ["top", "bottom", "left", "right"] as const) {
    const specific = get(ODF.fo, `border-${side}`), border = specific !== undefined ? parseBorder(specific) : all;
    if (border) borders[side] = edge(border);
  }
  const up = parseBorder(get(ODF.style, "diagonal-bl-tr")), down = parseBorder(get(ODF.style, "diagonal-tl-br"));
  if (up || down) { borders.diagonal = edge((up ?? down)!); if (up) borders.diagonalUp = true; if (down) borders.diagonalDown = true; }
  if (Object.keys(borders).length) style.borders = borders;

  const alignment: SpreadsheetAlignment = {};
  const textAlign = resolved.properties.get("paragraph-properties")?.get(key(ODF.fo, "text-align"));
  if (get(ODF.style, "text-align-source") !== "value-type" && textAlign !== undefined) {
    const horizontal = ({ start: "left", left: "left", end: "right", right: "right", center: "center", justify: "justify" } as const)[textAlign as "start"];
    if (horizontal) alignment.horizontal = horizontal;
  }
  const vertical = ({ top: "top", middle: "center", bottom: "bottom" } as const)[get(ODF.style, "vertical-align") as "top"];
  if (vertical) alignment.vertical = vertical;
  if (get(ODF.fo, "wrap-option") === "wrap") alignment.wrapText = true;
  if (get(ODF.style, "shrink-to-fit") === "true") alignment.shrinkToFit = true;
  const rotation = rotationOf(get(ODF.style, "rotation-angle"));
  if (rotation) alignment.textRotation = rotation;
  if (Object.keys(alignment).length) style.alignment = alignment;

  if (resolved.dataStyle !== undefined) {
    const format = numberStyles.get(resolved.dataStyle);
    if (format) style.numberFormatInfo = format.info;
  }
  return style;
}

/**
 * style:rotation-angle (degrees, optionally with "deg") as an Excel-style textRotation:
 * 1-90 counter-clockwise as is, 270-359 as 91-180 (clockwise). Upside-down angles are omitted.
 */
function rotationOf(value: string | undefined): number | undefined {
  const match = value === undefined ? null : /^([+-]?\d+(?:\.\d+)?)(deg)?$/.exec(value.trim());
  if (!match) return undefined;
  const angle = ((Math.round(Number(match[1])) % 360) + 360) % 360;
  if (angle === 0) return undefined;
  if (angle <= 90) return angle;
  if (angle >= 270) return 90 + (360 - angle);
  return undefined;
}

function localeOf(properties: Properties): string | undefined {
  const text = properties.get("text-properties");
  const language = text?.get(key(ODF.fo, "language")), country = text?.get(key(ODF.fo, "country"));
  if (!language || !/^[a-zA-Z]{2,3}$/.test(language) || language.toLowerCase() === "zxx") return undefined;
  return country && /^[a-zA-Z]{2}$/.test(country) ? `${language.toLowerCase()}-${country.toUpperCase()}` : language.toLowerCase();
}
