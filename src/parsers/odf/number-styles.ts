import type { SpreadsheetCellValue, SpreadsheetNumberFormatCategory, SpreadsheetNumberFormatInfo } from "../../model-extensions.js";
import { formatValue } from "../xlsx-number-format.js";
import { nsAttr, nsText, type NsXmlNode } from "../../xml/ns-tree.js";
import { EXT, ODF } from "./namespaces.js";
import { odfColor } from "./units.js";
import { dateSerial, DEFAULT_NULL_DATE, parseXsdDate } from "./values.js";

/**
 * OpenDocument data styles (ODF 1.3 Part 3 §16.29; implementation specification M13 and
 * decisions D9/D10). Internal module.
 *
 * Each data style is read into sections: the style itself plus the styles its style:map
 * elements apply. A section is converted to an Excel-compatible format code and formatted with
 * the shared `formatValue`, in the document's declared locale, never the viewer's (I-3).
 * Anything that cannot be reproduced exactly makes the section unsupported. Saved display text
 * is always what the grid shows (I-2); formatting is used only for cells without saved text and
 * to verify saved text.
 */

export type OdfNumberSection = {
  /** Excel-compatible code; absent when the section cannot be reproduced exactly. */
  code?: string;
  /** fo:color of the section's style:text-properties. */
  color?: string;
  /** BCP 47 tag used for separators and names. */
  locale: string;
  /** "date" when the section shows calendar fields, "time" for clock fields only. */
  dateTime: "none" | "date" | "time";
  /** The section is a number:boolean-style. */
  boolean?: true;
};
type Condition = { op: "<" | "<=" | ">" | ">=" | "=" | "!="; limit: number };
export type OdfNumberFormat = {
  name: string;
  category: SpreadsheetNumberFormatCategory;
  main: OdfNumberSection;
  maps: { condition: Condition; section: OdfNumberSection }[];
  /** Every section and condition can be reproduced. */
  supported: boolean;
  /** Every style:map condition could be evaluated, so the section for a value is known. */
  routable: boolean;
  info: SpreadsheetNumberFormatInfo;
};
export type OdfNumberStyles = { get(name: string): OdfNumberFormat | undefined };

const CATEGORIES: Readonly<Record<string, SpreadsheetNumberFormatCategory>> = {
  "number-style": "number", "currency-style": "currency", "percentage-style": "percentage",
  "date-style": "date", "time-style": "time", "boolean-style": "boolean", "text-style": "text",
};
const DATE_PARTS = new Set(["day", "month", "year", "day-of-week"]);
const TIME_PARTS = new Set(["hours", "minutes", "seconds", "am-pm"]);
const NUMBER_PARTS = new Set(["number", "scientific-number", "fraction"]);
/** Date fields formatValue cannot express. */
const UNSUPPORTED_PARTS = new Set(["era", "week-of-year", "quarter", "fill-character", "embedded-text"]);

/**
 * Data styles from `containers` (office:automatic-styles of content.xml first, then office:styles
 * of styles.xml: an earlier container wins a name clash). Styles are read on first use.
 */
export function createNumberStyles(containers: readonly (NsXmlNode | undefined)[], fallbackLocale: string | undefined): OdfNumberStyles {
  const nodes = new Map<string, NsXmlNode>();
  for (const container of containers) for (const child of container?.children ?? []) {
    if (typeof child === "string" || child.uri !== ODF.number || !Object.hasOwn(CATEGORIES, child.local)) continue;
    const name = nsAttr(child, ODF.style, "name");
    if (name !== undefined && !nodes.has(name)) nodes.set(name, child);
  }
  const defaultLocale = usableLocale(fallbackLocale) ?? "en-US";
  const cache = new Map<string, OdfNumberFormat | undefined>();
  return {
    get(name) {
      if (cache.has(name)) return cache.get(name);
      const node = nodes.get(name), format = node ? readFormat(name, node, nodes, defaultLocale) : undefined;
      cache.set(name, format);
      return format;
    },
  };
}

function readFormat(name: string, node: NsXmlNode, nodes: ReadonlyMap<string, NsXmlNode>, defaultLocale: string): OdfNumberFormat {
  const category = CATEGORIES[node.local]!;
  const main = readSection(node, defaultLocale);
  const maps: OdfNumberFormat["maps"] = [];
  let supported = main.section.code !== undefined || main.section.boolean === true, routable = true;
  for (const child of node.children) {
    if (typeof child === "string" || child.uri !== ODF.style || child.local !== "map") continue;
    const condition = parseCondition(nsAttr(child, ODF.style, "condition")), target = nsAttr(child, ODF.style, "apply-style-name");
    const targetNode = target === undefined ? undefined : nodes.get(target);
    // Conditions that cannot be evaluated make the whole style unusable: the section is unknown.
    if (!condition || !targetNode) { supported = false; routable = false; maps.length = 0; break; }
    const section = readSection(targetNode, main.section.locale).section;
    if (section.code === undefined && !section.boolean) supported = false;
    maps.push({ condition, section });
  }
  const code = supported ? summaryCode(main.section, maps) : undefined;
  const info: SpreadsheetNumberFormatInfo = {
    category, sourceName: name, supported,
    ...(code !== undefined ? { code } : {}),
    locale: main.section.locale,
    ...(main.currencySymbol !== undefined ? { currencySymbol: main.currencySymbol } : {}),
    ...(main.decimalPlaces !== undefined ? { decimalPlaces: main.decimalPlaces } : {}),
  };
  return { name, category, main: main.section, maps, supported, routable, info };
}

/** Informational Excel-style code: map sections with their conditions, then the style itself. */
function summaryCode(main: OdfNumberSection, maps: OdfNumberFormat["maps"]): string | undefined {
  if (main.boolean) return undefined;
  const parts = maps.map(({ condition, section }) => `[${condition.op === "!=" ? "<>" : condition.op}${condition.limit}]${section.code ?? ""}`);
  return [...parts, main.code ?? ""].join(";");
}

/** `value()>=0` and similar (ODF 1.3 Part 3 §19.473); other expressions are not evaluated. */
function parseCondition(text: string | undefined): Condition | undefined {
  const match = text === undefined ? null : /^\s*value\(\)\s*(<=|>=|!=|<>|<|>|=)\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)\s*$/.exec(text);
  if (!match) return undefined;
  const op = match[1] === "<>" ? "!=" : match[1] as Condition["op"], limit = Number(match[2]);
  return Number.isFinite(limit) ? { op, limit } : undefined;
}

function test(condition: Condition, value: number): boolean {
  const { op, limit } = condition;
  return op === "<" ? value < limit : op === "<=" ? value <= limit : op === ">" ? value > limit : op === ">=" ? value >= limit : op === "=" ? value === limit : value !== limit;
}

/** The section a value is shown with: the first map whose condition holds, else the style itself. */
function route(format: OdfNumberFormat, value: number): number {
  for (let index = 0; index < format.maps.length; index++) if (test(format.maps[index]!.condition, value)) return index;
  return -1;
}

/**
 * Whether no value >= 0 can reach section `index`. Such a section shows negative values without
 * their sign: a producer writes the sign as text (`-` or parentheses), like Excel's second section.
 * Routing is constant between condition limits, so the limits and points between them suffice.
 */
function negativeOnly(format: OdfNumberFormat, index: number): boolean {
  const limits = [...new Set([0, ...format.maps.map(map => map.condition.limit).filter(limit => limit >= 0)])].sort((a, b) => a - b);
  const points = [...limits, limits[limits.length - 1]! * 2 + 1];
  for (let i = 0; i + 1 < limits.length; i++) points.push((limits[i]! + limits[i + 1]!) / 2);
  return points.every(point => route(format, point) !== index);
}

type SectionRead = { section: OdfNumberSection; currencySymbol?: string; decimalPlaces?: number };

/** Reads one data style element into a section (its own style:map children are not followed). */
function readSection(node: NsXmlNode, inheritedLocale: string): SectionRead {
  const declared = declaredLocale(node);
  const locale = declared === undefined ? inheritedLocale : usableLocale(declared);
  const category = CATEGORIES[node.local]!;
  const parts: string[] = [];
  const kinds: { part: string; long: boolean; textual?: boolean }[] = [];
  let ok = locale !== undefined, color: string | undefined, currencySymbol: string | undefined, decimalPlaces: number | undefined;
  let numeric = 0, boolean = false, text = false, dateParts = 0, timeParts = 0;
  const unsupported = () => { ok = false; };

  // Attributes that change the output in ways a format code cannot express.
  if (nsAttr(node, ODF.number, "automatic-order") === "true" || nsAttr(node, ODF.number, "format-source") === "language") unsupported();
  if (node.attributes.some(attribute => attribute.uri === ODF.number && attribute.local.startsWith("transliteration-"))) unsupported();
  const elapsed = node.local === "time-style" && nsAttr(node, ODF.number, "truncate-on-overflow") === "false";
  let elapsedPending = elapsed;

  for (const child of node.children) {
    if (typeof child === "string") continue;
    if (child.uri === ODF.style && child.local === "text-properties") { color = odfColor(nsAttr(child, ODF.fo, "color")); continue; }
    if (child.uri === ODF.style && child.local === "map") continue;
    if (child.uri !== ODF.number || UNSUPPORTED_PARTS.has(child.local)) { unsupported(); continue; }
    const calendar = nsAttr(child, ODF.number, "calendar");
    if (calendar !== undefined && calendar !== "gregorian") unsupported();
    const long = nsAttr(child, ODF.number, "style") === "long";
    switch (child.local) {
      case "text": parts.push(category === "percentage" ? percentLiteral(nsText(child)) : literal(nsText(child))); break;
      case "currency-symbol": {
        const symbol = nsText(child);
        if (!symbol) unsupported(); // an empty symbol means the locale's default, which is not known here
        currencySymbol ??= symbol;
        parts.push(literal(symbol));
        break;
      }
      case "number": case "scientific-number": case "fraction": {
        numeric++;
        const read = child.local === "number" ? numberCode(child) : child.local === "scientific-number" ? scientificCode(child) : fractionCode(child);
        if (read === undefined) unsupported(); else parts.push(read);
        if (child.local !== "fraction") { const places = integerAttribute(child, ODF.number, "decimal-places"); if (places !== undefined) decimalPlaces ??= places; }
        break;
      }
      case "boolean": boolean = true; break;
      case "text-content": text = true; parts.push("@"); break;
      case "day": dateParts++; kinds.push({ part: "day", long }); parts.push(long ? "dd" : "d"); break;
      case "month": {
        dateParts++;
        const textual = nsAttr(child, ODF.number, "textual") === "true";
        if (nsAttr(child, ODF.number, "possessive-form") === "true") unsupported();
        kinds.push({ part: "month", long, textual });
        parts.push(textual ? (long ? "mmmm" : "mmm") : (long ? "mm" : "m"));
        break;
      }
      case "year": dateParts++; kinds.push({ part: "year", long }); parts.push(long ? "yyyy" : "yy"); break;
      case "day-of-week": dateParts++; kinds.push({ part: "day-of-week", long }); parts.push(long ? "dddd" : "ddd"); break;
      case "hours": case "minutes": case "seconds": {
        timeParts++;
        const letter = child.local === "hours" ? "h" : child.local === "minutes" ? "m" : "s";
        let token = long ? letter + letter : letter;
        if (elapsedPending) { token = `[${token}]`; elapsedPending = false; }
        const places = child.local === "seconds" ? integerAttribute(child, ODF.number, "decimal-places") ?? 0 : 0;
        if (places > 9) unsupported();
        kinds.push({ part: child.local, long });
        parts.push(token + (places ? "." + "0".repeat(places) : ""));
        break;
      }
      case "am-pm": timeParts++; parts.push("AM/PM"); break;
      default: unsupported();
    }
  }

  // Shape checks per category: one number element for numeric styles, fields only for dates.
  const isNumeric = category === "number" || category === "currency" || category === "percentage";
  const others = dateParts || timeParts || boolean || text;
  // One number element, or text only (e.g. a "-" section that a style:map applies to zero).
  if (isNumeric && (others || numeric > 1 || (numeric === 0 && !parts.length))) unsupported();
  if (category === "percentage" && numeric && !parts.some(part => part.replace(/"[^"]*"/g, "").includes("%"))) unsupported();
  if ((category === "date" || category === "time") && (numeric || boolean || text || !(dateParts || timeParts))) unsupported();
  if (category === "boolean" && (!boolean || numeric || dateParts || timeParts || parts.length)) unsupported();
  // Text styles show strings as saved; numbers under them use the producer's General format.
  if (category === "text") unsupported();
  if (elapsedPending) unsupported(); // truncate-on-overflow="false" without a clock field to apply it to
  if (!ambiguitySafe(kinds)) unsupported();

  const dateTime = dateParts ? "date" : timeParts ? "time" : "none";
  const section: OdfNumberSection = { locale: locale ?? inheritedLocale, dateTime, ...(color ? { color } : {}) };
  if (boolean && ok) section.boolean = true;
  else if (ok) section.code = parts.join("");
  return { section, ...(currencySymbol !== undefined ? { currencySymbol } : {}), ...(decimalPlaces !== undefined ? { decimalPlaces } : {}) };
}

/**
 * formatValue reads "m"/"mm" as minutes right after hours or right before seconds, otherwise as
 * months (Excel's rule). Sections where ODF's explicit month/minute elements would be read the
 * other way are not reproducible.
 */
function ambiguitySafe(kinds: readonly { part: string; textual?: boolean }[]): boolean {
  const clock = kinds.filter(kind => kind.part !== "day-of-week" && !(kind.part === "month" && kind.textual));
  return clock.every((kind, index) => {
    const previous = clock[index - 1]?.part, next = clock[index + 1]?.part;
    const readAsMinutes = previous === "hours" || next === "seconds";
    if (kind.part === "minutes") return readAsMinutes;
    if (kind.part === "month") return !readAsMinutes;
    return true;
  });
}

/** number:number as an Excel code; undefined for General (no decimal-places) and inexact options. */
function numberCode(node: NsXmlNode): string | undefined {
  if (node.children.some(child => typeof child !== "string")) return undefined; // number:embedded-text
  if (nsAttr(node, ODF.number, "decimal-replacement") !== undefined) return undefined;
  const decimals = integerAttribute(node, ODF.number, "decimal-places");
  // Without decimal-places the producer's "General" applies, whose text depends on the column width.
  if (decimals === undefined) return undefined;
  const minDecimals = integerAttribute(node, ODF.number, "min-decimal-places") ?? integerAttribute(node, EXT.loext, "min-decimal-places") ?? decimals;
  // Optional decimals ("0.##") end differently in producers when no decimal is shown.
  if (minDecimals !== decimals) return undefined;
  const minInteger = integerAttribute(node, ODF.number, "min-integer-digits") ?? 0;
  if (minInteger > 30 || decimals > 30) return undefined;
  const factor = displayFactorCommas(nsAttr(node, ODF.number, "display-factor"));
  if (factor === undefined) return undefined;
  return integerCode(minInteger, nsAttr(node, ODF.number, "grouping") === "true") + factor + (decimals ? "." + "0".repeat(decimals) : "");
}

function scientificCode(node: NsXmlNode): string | undefined {
  if (node.children.some(child => typeof child !== "string") || nsAttr(node, ODF.number, "grouping") === "true") return undefined;
  const decimals = integerAttribute(node, ODF.number, "decimal-places") ?? 0;
  const minDecimals = integerAttribute(node, ODF.number, "min-decimal-places") ?? integerAttribute(node, EXT.loext, "min-decimal-places") ?? decimals;
  if (minDecimals !== decimals || decimals > 30) return undefined;
  const minInteger = integerAttribute(node, ODF.number, "min-integer-digits") ?? 1;
  const exponentDigits = integerAttribute(node, ODF.number, "min-exponent-digits") ?? 2;
  const interval = integerAttribute(node, ODF.number, "exponent-interval") ?? integerAttribute(node, EXT.loext, "exponent-interval") ?? 1;
  const forced = (nsAttr(node, ODF.number, "forced-exponent-sign") ?? nsAttr(node, EXT.loext, "forced-exponent-sign")) !== "false";
  if (minInteger < 1 || minInteger > 30 || exponentDigits < 1 || exponentDigits > 9 || interval < 1) return undefined;
  // Engineering notation: formatValue keeps the exponent a multiple of the "#"-padded integer width.
  if (interval > 1 && minInteger >= interval) return undefined;
  const integer = interval > 1 ? "#".repeat(interval - minInteger) + "0".repeat(minInteger) : "0".repeat(minInteger);
  return integer + (decimals ? "." + "0".repeat(decimals) : "") + (forced ? "E+" : "E-") + "0".repeat(exponentDigits);
}

function fractionCode(node: NsXmlNode): string | undefined {
  if (node.children.some(child => typeof child !== "string") || nsAttr(node, ODF.number, "grouping") === "true") return undefined;
  const minInteger = integerAttribute(node, ODF.number, "min-integer-digits");
  const numerator = integerAttribute(node, ODF.number, "min-numerator-digits") ?? 1;
  const denominatorDigits = integerAttribute(node, ODF.number, "min-denominator-digits") ?? 1;
  const denominator = integerAttribute(node, ODF.number, "denominator-value");
  const maxDenominator = integerAttribute(node, ODF.number, "max-denominator-value") ?? integerAttribute(node, EXT.loext, "max-denominator-value");
  if (numerator < 1 || numerator > 9 || denominatorDigits < 1 || denominatorDigits > 6 || (minInteger !== undefined && minInteger > 30)) return undefined;
  // A fixed denominator ("# ?/16") is not read as a fraction by formatValue's code parser.
  if (denominator !== undefined) return undefined;
  if (maxDenominator !== undefined && maxDenominator !== 10 ** denominatorDigits - 1) return undefined;
  const integer = minInteger === undefined ? "" : (minInteger === 0 ? "#" : "0".repeat(minInteger)) + " ";
  return integer + "?".repeat(numerator) + "/" + "?".repeat(denominatorDigits);
}

function integerCode(minInteger: number, grouping: boolean): string {
  let pattern = "0".repeat(minInteger);
  if (grouping) {
    pattern = "#".repeat(Math.max(0, 4 - minInteger)) + pattern;
    return pattern.slice(0, -3) + "," + pattern.slice(-3);
  }
  return pattern || "#";
}

/** number:display-factor as Excel scaling commas: only powers of 1000 are exact. */
function displayFactorCommas(value: string | undefined): string | undefined {
  if (value === undefined) return "";
  const factor = Number(value);
  for (let commas = 0, scale = 1; commas <= 4; commas++, scale *= 1000) if (factor === scale) return ",".repeat(commas);
  return undefined;
}

function integerAttribute(node: NsXmlNode, uri: string, local: string): number | undefined {
  const value = nsAttr(node, uri, local);
  if (value === undefined || !/^\d+$/.test(value.trim())) return undefined;
  const number = Number(value);
  return Number.isSafeInteger(number) ? number : undefined;
}

/** A quoted Excel literal; a double quote is written escaped. */
function literal(text: string): string {
  return text.split('"').map(piece => piece ? `"${piece}"` : "").join('\\"');
}

/** Percentage-style text: "%" signs scale the value and are kept unquoted. */
function percentLiteral(text: string): string {
  return text.split("%").map(literal).join("%");
}

function declaredLocale(node: NsXmlNode): string | undefined {
  const tag = nsAttr(node, ODF.number, "rfc-language-tag");
  if (tag) return tag;
  const language = nsAttr(node, ODF.number, "language"), country = nsAttr(node, ODF.number, "country");
  if (!language) return undefined;
  return country ? `${language}-${country}` : language;
}

/** A locale the formatter's Intl data supports, canonicalised; undefined otherwise. */
function usableLocale(tag: string | undefined): string | undefined {
  if (!tag) return undefined;
  try {
    const [supported] = Intl.NumberFormat.supportedLocalesOf(tag);
    return supported && supported.toLowerCase().split("-")[0] === tag.toLowerCase().split("-")[0] ? supported : undefined;
  } catch { return undefined; }
}

export type OdfFormatted = { text?: string; color?: string };

/**
 * Formats a typed value with a data style. `text` is absent when the chosen section, or this
 * value with it, cannot be reproduced; `color` is the chosen section's colour. Undefined when no
 * section can be chosen (string, empty, error and invalid values, or unknown dates).
 */
export function formatOdfValue(format: OdfNumberFormat, value: SpreadsheetCellValue, nullDate: string): OdfFormatted | undefined {
  if (value.kind === "boolean") {
    const section = format.main;
    return section.boolean ? { text: formatValue(value.boolean, undefined, { locale: section.locale }).text, ...(section.color ? { color: section.color } : {}) } : undefined;
  }
  const number = value.kind === "number" || value.kind === "percentage" || value.kind === "currency" ? value.number
    : value.kind === "date" ? value.serial : value.kind === "time" ? value.days : undefined;
  if (number === undefined || !Number.isFinite(number) || !format.routable) return undefined;
  const index = route(format, number);
  const section = index < 0 ? format.main : format.maps[index]!.section;
  const color = section.color ? { color: section.color } : {};
  if (section.code === undefined) return color;
  let input = number;
  if (section.dateTime === "date") {
    // formatValue counts days in Excel's 1900 system, which is exact from 1900-03-01 (serial 61).
    const offset = nullDateOffset(nullDate);
    if (offset === undefined) return color;
    input = number + offset;
    if (input < 61) return color;
  } else if (section.dateTime === "time") {
    if (number < 0) return color;
  } else if (number < 0 && negativeOnly(format, index)) input = -number;
  return { text: formatValue(input, section.code, { locale: section.locale }).text, ...color };
}

const offsets = new Map<string, number | undefined>();
/** Days from 1899-12-30 (Excel's day 0 for serials from 61) to the document's null date. */
function nullDateOffset(nullDate: string): number | undefined {
  if (offsets.has(nullDate)) return offsets.get(nullDate);
  const parsed = parseXsdDate(nullDate);
  const offset = parsed && !parsed.hasTimezone && parsed.seconds === 0 ? dateSerial(parsed, DEFAULT_NULL_DATE) : undefined;
  offsets.set(nullDate, offset);
  return offset;
}

const SPACES = new RegExp(`[${String.fromCharCode(0xa0)}${String.fromCharCode(0x202f)}]`, "g");
const MINUS = new RegExp(String.fromCharCode(0x2212), "g");
const APOSTROPHE = new RegExp(String.fromCharCode(0x2019), "g");
/**
 * Whether saved and computed text agree once space and minus variants are unified (specification
 * §6.3 step 3). The typographic apostrophe is also unified: Swiss grouping is written either way,
 * depending on the producer's and the browser's locale data.
 */
export function sameDisplayText(saved: string, computed: string): boolean {
  const normalise = (text: string) => text.replace(SPACES, " ").replace(MINUS, "-").replace(APOSTROPHE, "'").trim();
  return normalise(saved) === normalise(computed);
}
