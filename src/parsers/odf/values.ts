import type { DiagnosticCollector } from "../../diagnostics.js";
import type { SpreadsheetCellValue } from "../../model-extensions.js";
import { nsAttr, type NsXmlNode } from "../../xml/ns-tree.js";
import { EXT, ODF } from "./namespaces.js";

/**
 * Typed OpenDocument cell values (ODF 1.3 Part 3 §19.385-19.391 and XML Schema datatypes;
 * implementation specification §6.2 and financial invariants I-1, I-4, I-8, I-9, I-12).
 *
 * Internal module. Values come only from the value attributes, never from display text. The
 * stored text is kept verbatim as `lexical`; anything that does not match its declared type
 * becomes `invalid` and is reported, never coerced.
 */

const doublePattern = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;
const datePattern = /^(-?\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2})(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?$/;
const durationPattern = /^(-)?P(?:(\d+)Y)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/;
export const DEFAULT_NULL_DATE = "1899-12-30";

/** An xsd:double as a finite number; INF, -INF, NaN and anything malformed are undefined. */
export function parseXsdDouble(lexical: string): number | undefined {
  const text = lexical.trim();
  if (!doublePattern.test(text)) return undefined;
  const value = Number(text);
  return Number.isFinite(value) ? value : undefined;
}

export type XsdDate = { year: number; month: number; day: number; seconds: number; hasTimezone: boolean };

/** An xsd:date or xsd:dateTime in the proleptic Gregorian calendar; `seconds` is the time of day. */
export function parseXsdDate(lexical: string): XsdDate | undefined {
  const match = datePattern.exec(lexical.trim());
  if (!match) return undefined;
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]);
  if (month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month)) return undefined;
  let seconds = 0;
  if (match[4] !== undefined) {
    const hours = Number(match[4]), minutes = Number(match[5]), whole = Number(match[6]), fraction = match[7] ? Number(match[7]) : 0;
    if (hours === 24 ? minutes !== 0 || whole !== 0 || fraction !== 0 : hours > 23 || minutes > 59 || whole > 59) return undefined;
    seconds = hours * 3600 + minutes * 60 + whole + fraction;
  }
  if (match[8] && match[8] !== "Z") {
    const [hours, minutes] = match[8].slice(1).split(":").map(Number);
    if (hours! > 14 || minutes! > 59) return undefined;
  }
  return { year, month, day, seconds, hasTimezone: match[8] !== undefined };
}

export type XsdDuration = { negative: boolean; days?: number; hasYearOrMonth: boolean };

/** An xsd:duration. `days` is the length in days, absent when years or months make it calendar-dependent. */
export function parseXsdDuration(lexical: string): XsdDuration | undefined {
  const text = lexical.trim(), match = durationPattern.exec(text);
  if (!match || text.endsWith("P") || text.endsWith("T") || match.slice(2).every(part => part === undefined)) return undefined;
  const negative = match[1] === "-", hasYearOrMonth = match[2] !== undefined || match[3] !== undefined;
  if (hasYearOrMonth) return { negative, hasYearOrMonth };
  const [days, hours, minutes, seconds] = [match[4], match[5], match[6], match[7]].map(part => part === undefined ? 0 : Number(part));
  const total = days! + hours! / 24 + minutes! / 1440 + seconds! / 86400;
  if (!Number.isFinite(total)) return undefined;
  return { negative, days: negative && total ? -total : total, hasYearOrMonth };
}

/**
 * Days (with the time of day as a fraction) from `nullDate`, an xsd:date such as "1899-12-30".
 * Integer calendar arithmetic: no time zones, no daylight saving, and no phantom 29 February 1900.
 */
export function dateSerial(date: Omit<XsdDate, "hasTimezone">, nullDate: string): number {
  const epoch = parseXsdDate(nullDate);
  if (!epoch) throw new RangeError(`Invalid null date ${nullDate}.`);
  return daysFromCivil(date.year, date.month, date.day) - daysFromCivil(epoch.year, epoch.month, epoch.day) + date.seconds / 86400;
}

/**
 * The typed value of a table:table-cell (specification §6.2). `savedText` is the cell's saved
 * display text (savedCellText), used only for string and error cells, never to derive numbers.
 */
export function readCellValue(cell: NsXmlNode, savedText: string | undefined, collector: DiagnosticCollector, nullDate: string, location: { sheet?: string; cell?: string } = {}): SpreadsheetCellValue {
  if (nsAttr(cell, EXT.calcext, "value-type") === "error") return { kind: "error", code: savedText ?? "#ERROR" };
  const declaredType = nsAttr(cell, ODF.office, "value-type");
  const invalid = (lexical?: string): SpreadsheetCellValue => {
    collector.report("odf.value.invalid", location);
    return { kind: "invalid", declaredType: declaredType ?? "", ...(lexical !== undefined ? { lexical } : {}) };
  };
  switch (declaredType) {
    case undefined: return savedText !== undefined ? { kind: "string", text: savedText } : { kind: "empty" };
    case "void": return { kind: "empty" };
    case "float": case "percentage": case "currency": {
      const lexical = nsAttr(cell, ODF.office, "value");
      if (lexical === undefined) return invalid();
      const number = parseXsdDouble(lexical);
      if (number === undefined) return invalid(lexical);
      if (declaredType !== "currency") return { kind: declaredType === "float" ? "number" : "percentage", lexical, number };
      const currency = nsAttr(cell, ODF.office, "currency");
      return { kind: "currency", lexical, number, ...(currency !== undefined ? { currency } : {}) };
    }
    case "date": {
      const lexical = nsAttr(cell, ODF.office, "date-value");
      if (lexical === undefined) return invalid();
      const date = parseXsdDate(lexical);
      if (!date) return invalid(lexical);
      if (date.hasTimezone) collector.report("odf.value.timezone", location);
      return parseXsdDate(nullDate) ? { kind: "date", lexical, serial: dateSerial(date, nullDate) } : { kind: "date", lexical };
    }
    case "time": {
      const lexical = nsAttr(cell, ODF.office, "time-value");
      if (lexical === undefined) return invalid();
      const duration = parseXsdDuration(lexical);
      if (!duration) return invalid(lexical);
      return duration.days === undefined ? { kind: "time", lexical } : { kind: "time", lexical, days: duration.days };
    }
    case "boolean": {
      const lexical = nsAttr(cell, ODF.office, "boolean-value");
      if (lexical === undefined) return invalid();
      const text = lexical.trim();
      if (text === "true" || text === "1") return { kind: "boolean", lexical, boolean: true };
      if (text === "false" || text === "0") return { kind: "boolean", lexical, boolean: false };
      return invalid(lexical);
    }
    case "string": return { kind: "string", text: nsAttr(cell, ODF.office, "string-value") ?? savedText ?? "" };
    default: return invalid();
  }
}

function isLeapYear(year: number): boolean { return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0; }
function daysInMonth(year: number, month: number): number { return month === 2 ? (isLeapYear(year) ? 29 : 28) : [4, 6, 9, 11].includes(month) ? 30 : 31; }

/** Days since 1970-01-01 in the proleptic Gregorian calendar (H. Hinnant's days_from_civil). */
function daysFromCivil(year: number, month: number, day: number): number {
  const y = year - (month <= 2 ? 1 : 0), era = Math.floor(y / 400), yearOfEra = y - era * 400;
  const dayOfYear = Math.floor((153 * (month + (month > 2 ? -3 : 9)) + 2) / 5) + day - 1;
  const dayOfEra = yearOfEra * 365 + Math.floor(yearOfEra / 4) - Math.floor(yearOfEra / 100) + dayOfYear;
  return era * 146_097 + dayOfEra - 719_468;
}
