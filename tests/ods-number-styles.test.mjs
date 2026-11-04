// ODF data styles and saved-text verification (implementation specification M13, decisions D9/D10).
import test from "node:test";
import assert from "node:assert/strict";
import { createNumberStyles, formatOdfValue, sameDisplayText } from "../dist/parsers/odf/number-styles.js";
import { parseOds } from "../dist/parsers/odf/ods.js";
import { parseXmlNs } from "../dist/xml/ns-tree.js";
import { makeOdfPackage, makeOds, odfXml, spreadsheetContent, spreadsheetStyles } from "./helpers/odf.mjs";

const EURO = String.fromCharCode(0x20ac), YEN = String.fromCharCode(0xa5), NBSP = String.fromCharCode(0xa0), NNBSP = String.fromCharCode(0x202f), MINUS = String.fromCharCode(0x2212);
const PREFIX_SETS = [["canonical prefixes", {}], ["renamed prefixes", { office: "o", table: "t", text: "x", style: "s", fo: "f", number: "n", calcext: "c", svg: "g", meta: "m", manifest: "mf", of: "q", loext: "lo" }]];
const NULL_DATE = "1899-12-30";

/** Data styles in an office:automatic-styles container; `fallback` is the document locale. */
function stylesOf(definitions, prefixes = {}, fallback = "en-US") {
  return createNumberStyles([parseXmlNs(odfXml(`<office:automatic-styles>${definitions}</office:automatic-styles>`, prefixes))], fallback);
}
const num = (decimals, extra = "") => `<number:number number:decimal-places="${decimals}" number:min-decimal-places="${decimals}" number:min-integer-digits="1"${extra}/>`;
const grouped = (decimals, extra = "") => num(decimals, ` number:grouping="true"${extra}`);
const text = value => `<number:text>${value}</number:text>`;
const lang = (language, country) => ` number:language="${language}" number:country="${country}"`;
const n = value => ({ kind: "number", lexical: String(value), number: value });
const pct = value => ({ kind: "percentage", lexical: String(value), number: value });
const cur = (value, currency) => ({ kind: "currency", lexical: String(value), number: value, ...(currency ? { currency } : {}) });
// Date serials relative to 1899-12-30, as values.ts computes them (2026-10-28 = 46323).
const date = (lexical, serial) => ({ kind: "date", lexical, serial });
const time = (lexical, days) => ({ kind: "time", lexical, days });

/** [label, definition (the style named "S"), expected code, value, expected text] */
const CASES = [
  ["fixed decimals", `<number:number-style style:name="S">${num(2)}</number:number-style>`, "0.00", n(1234.5), "1234.50"],
  ["grouping with decimals", `<number:number-style style:name="S">${grouped(2)}</number:number-style>`, "#,##0.00", n(1234.5), "1,234.50"],
  ["grouping without decimals", `<number:number-style style:name="S">${grouped(0)}</number:number-style>`, "#,##0", n(1234567), "1,234,567"],
  ["minimum integer digits", `<number:number-style style:name="S"><number:number number:decimal-places="0" number:min-integer-digits="3"/></number:number-style>`, "000", n(7), "007"],
  ["no integer digit", `<number:number-style style:name="S"><number:number number:decimal-places="2" number:min-integer-digits="0"/></number:number-style>`, "#.00", n(0.5), ".50"],
  ["negative sign kept without a map", `<number:number-style style:name="S">${grouped(2)}</number:number-style>`, "#,##0.00", n(-1234.5), "-1,234.50"],
  ["display factor 1000", `<number:number-style style:name="S"><number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true" number:display-factor="1000"/></number:number-style>`, "#,##0,", n(1234567), "1,235"],
  ["display factor 1000000", `<number:number-style style:name="S"><number:number number:decimal-places="1" number:min-integer-digits="1" number:display-factor="1000000"/></number:number-style>`, "0,,.0", n(2500000), "2.5"],
  ["unit suffix", `<number:number-style style:name="S">${num(1)}${text(" kg")}</number:number-style>`, '0.0" kg"', n(12.5), "12.5 kg"],
  ["text with a double quote", `<number:number-style style:name="S">${num(0)}${text('"')}</number:number-style>`, '0\\"', n(12), '12"'],
  ["currency after with a space (de-DE)", `<number:currency-style style:name="S"${lang("de", "DE")}>${grouped(2)}${text(" ")}<number:currency-symbol${lang("de", "DE")}>${EURO}</number:currency-symbol></number:currency-style>`, `#,##0.00" ""${EURO}"`, cur(1234.5, "EUR"), `1.234,50 ${EURO}`],
  ["currency before (en-US)", `<number:currency-style style:name="S"${lang("en", "US")}><number:currency-symbol${lang("en", "US")}>$</number:currency-symbol>${grouped(2)}</number:currency-style>`, '"$"#,##0.00', cur(1234.5, "USD"), "$1,234.50"],
  ["currency before with a space (en-GB)", `<number:currency-style style:name="S"${lang("en", "GB")}><number:currency-symbol>CHF</number:currency-symbol>${text(" ")}${grouped(2)}</number:currency-style>`, '"CHF"" "#,##0.00', cur(1234.5, "CHF"), "CHF 1,234.50"],
  ["yen without decimals (ja-JP)", `<number:currency-style style:name="S"${lang("ja", "JP")}><number:currency-symbol>${YEN}</number:currency-symbol>${grouped(0)}</number:currency-style>`, `"${YEN}"#,##0`, cur(1235, "JPY"), `${YEN}1,235`],
  ["percentage", `<number:percentage-style style:name="S">${num(0)}${text("%")}</number:percentage-style>`, "0%", pct(0.15), "15%"],
  ["percentage with decimals", `<number:percentage-style style:name="S">${num(2)}${text("%")}</number:percentage-style>`, "0.00%", pct(0.1234), "12.34%"],
  ["percentage with a space (fr-FR)", `<number:percentage-style style:name="S"${lang("fr", "FR")}>${num(1)}${text(" %")}</number:percentage-style>`, '0.0" "%', pct(0.155), "15,5 %"],
  ["scientific", `<number:number-style style:name="S"><number:scientific-number number:decimal-places="2" number:min-integer-digits="1" number:min-exponent-digits="2"/></number:number-style>`, "0.00E+00", n(12345), "1.23E+04"],
  ["engineering notation", `<number:number-style style:name="S"><number:scientific-number number:decimal-places="2" number:min-integer-digits="1" number:min-exponent-digits="2" number:exponent-interval="3"/></number:number-style>`, "##0.00E+00", n(12340), "12.34E+03"],
  ["mixed fraction", `<number:number-style style:name="S"><number:fraction number:min-integer-digits="0" number:min-numerator-digits="1" number:min-denominator-digits="1"/></number:number-style>`, "# ?/?", n(1.5), "1 1/2"],
  ["improper fraction", `<number:number-style style:name="S"><number:fraction number:min-numerator-digits="1" number:min-denominator-digits="1"/></number:number-style>`, "?/?", n(0.75), "3/4"],
  ["ISO date", `<number:date-style style:name="S"><number:year number:style="long"/>${text("-")}<number:month number:style="long"/>${text("-")}<number:day number:style="long"/></number:date-style>`, 'yyyy"-"mm"-"dd', date("2026-10-28", 46323), "2026-10-28"],
  ["long date (en-US)", `<number:date-style style:name="S"${lang("en", "US")}><number:day-of-week number:style="long"/>${text(", ")}<number:month number:style="long" number:textual="true"/>${text(" ")}<number:day/>${text(", ")}<number:year number:style="long"/></number:date-style>`, 'dddd", "mmmm" "d", "yyyy', date("2026-10-28", 46323), "Wednesday, October 28, 2026"],
  ["short date (de-DE)", `<number:date-style style:name="S"${lang("de", "DE")}><number:day number:style="long"/>${text(".")}<number:month number:style="long"/>${text(".")}<number:year number:style="long"/></number:date-style>`, 'dd"."mm"."yyyy', date("2026-10-28", 46323), "28.10.2026"],
  ["abbreviated month", `<number:date-style style:name="S"><number:day/>${text(" ")}<number:month number:textual="true"/>${text(" ")}<number:year number:style="long"/></number:date-style>`, 'd" "mmm" "yyyy', date("2026-10-28", 46323), "28 Oct 2026"],
  ["two-digit year", `<number:date-style style:name="S"><number:month/>${text("/")}<number:day/>${text("/")}<number:year/></number:date-style>`, 'm"/"d"/"yy', date("2026-03-05", 46086), "3/5/26"],
  ["date and time", `<number:date-style style:name="S"><number:day number:style="long"/>${text(".")}<number:month number:style="long"/>${text(".")}<number:year number:style="long"/>${text(" ")}<number:hours number:style="long"/>${text(":")}<number:minutes number:style="long"/></number:date-style>`, 'dd"."mm"."yyyy" "hh":"mm', date("2026-10-28T13:05:00", 46323 + (13 * 60 + 5) / 1440), "28.10.2026 13:05"],
  ["time of day", `<number:time-style style:name="S"><number:hours number:style="long"/>${text(":")}<number:minutes number:style="long"/>${text(":")}<number:seconds number:style="long"/></number:time-style>`, 'hh":"mm":"ss', time("PT13H05M09S", (13 * 3600 + 5 * 60 + 9) / 86400), "13:05:09"],
  ["12-hour clock", `<number:time-style style:name="S"><number:hours/>${text(":")}<number:minutes number:style="long"/>${text(" ")}<number:am-pm/></number:time-style>`, 'h":"mm" "AM/PM', time("PT13H05M00S", (13 * 60 + 5) / 1440), "1:05 PM"],
  ["elapsed hours", `<number:time-style style:name="S" number:truncate-on-overflow="false"><number:hours number:style="long"/>${text(":")}<number:minutes number:style="long"/>${text(":")}<number:seconds number:style="long"/></number:time-style>`, '[hh]":"mm":"ss', time("PT36H30M00S", 36.5 / 24), "36:30:00"],
  ["hours truncated by default", `<number:time-style style:name="S"><number:hours number:style="long"/>${text(":")}<number:minutes number:style="long"/></number:time-style>`, 'hh":"mm', time("PT36H30M00S", 36.5 / 24), "12:30"],
  ["fractional seconds", `<number:time-style style:name="S"><number:minutes number:style="long"/>${text(":")}<number:seconds number:style="long" number:decimal-places="2"/></number:time-style>`, 'mm":"ss.00', time("PT0H05M09.5S", 309.5 / 86400), "05:09.50"],
  ["time style on a date-time value", `<number:time-style style:name="S"><number:hours number:style="long"/>${text(":")}<number:minutes number:style="long"/></number:time-style>`, 'hh":"mm', date("2026-10-28T08:15:00", 46323 + (8 * 60 + 15) / 1440), "08:15"],
];

test("data styles convert to Excel codes and format in the document locale (30+ definitions)", () => {
  assert.ok(CASES.length >= 30);
  for (const [prefixLabel, prefixes] of PREFIX_SETS) {
    for (const [label, definition, code, value, expected] of CASES) {
      const format = stylesOf(definition, prefixes).get("S");
      assert.ok(format, `${prefixLabel} ${label}: read`);
      assert.equal(format.supported, true, `${prefixLabel} ${label}: supported`);
      assert.equal(format.info.code, code, `${prefixLabel} ${label}: code`);
      const formatted = formatOdfValue(format, value, NULL_DATE);
      assert.equal(formatted?.text?.trim(), expected, `${prefixLabel} ${label}: text`);
    }
  }
});

test("format info records category, locale, currency symbol and decimals", () => {
  const styles = stylesOf(`<number:currency-style style:name="C"${lang("de", "DE")}>${grouped(2)}${text(" ")}<number:currency-symbol>${EURO}</number:currency-symbol></number:currency-style><number:number-style style:name="N">${num(3)}</number:number-style>`, {}, "fr-CH");
  assert.deepEqual(styles.get("C").info, { category: "currency", sourceName: "C", supported: true, code: `#,##0.00" ""${EURO}"`, locale: "de-DE", currencySymbol: EURO, decimalPlaces: 2 });
  assert.equal(styles.get("N").info.locale, "fr-CH", "no number:language: the document locale");
  assert.equal(stylesOf(`<number:number-style style:name="N">${num(0)}</number:number-style>`, {}, undefined).get("N").info.locale, "en-US", "no locale anywhere: en-US");
  assert.equal(styles.get("missing"), undefined);
});

test("separators come from the document locale, never the viewer's (de-DE, fr-CH)", () => {
  const value = n(1234567.891);
  const de = stylesOf(`<number:number-style style:name="S"${lang("de", "DE")}>${grouped(2)}</number:number-style>`).get("S");
  assert.equal(formatOdfValue(de, value, NULL_DATE).text, "1.234.567,89");
  const frCH = stylesOf(`<number:number-style style:name="S"${lang("fr", "CH")}>${grouped(2)}</number:number-style>`).get("S");
  const expected = new Intl.NumberFormat("fr-CH", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(1234567.891);
  assert.ok(sameDisplayText(formatOdfValue(frCH, value, NULL_DATE).text, expected), "fr-CH grouping and decimal separators");
  const inherited = stylesOf(`<number:number-style style:name="S">${grouped(2)}</number:number-style>`, {}, "de-DE").get("S");
  assert.equal(formatOdfValue(inherited, value, NULL_DATE).text, "1.234.567,89", "the document default locale");
  assert.equal(stylesOf(`<number:boolean-style style:name="B"${lang("de", "DE")}><number:boolean/></number:boolean-style>`).get("B") && formatOdfValue(stylesOf(`<number:boolean-style style:name="B"${lang("de", "DE")}><number:boolean/></number:boolean-style>`).get("B"), { kind: "boolean", lexical: "true", boolean: true }, NULL_DATE).text, "WAHR");
  assert.equal(formatOdfValue(stylesOf(`<number:boolean-style style:name="B"><number:boolean/></number:boolean-style>`).get("B"), { kind: "boolean", lexical: "false", boolean: false }, NULL_DATE).text, "FALSE");
});

// LibreOffice writes signed formats as a style for the "otherwise" case plus style:map elements.
const positive = `<number:number-style style:name="P0" style:volatile="true">${grouped(2)}</number:number-style>`;
const negativeRed = `${positive}<number:number-style style:name="S"><style:text-properties fo:color="#FF0000"/>${text("-")}${grouped(2)}<style:map style:condition="value()&gt;=0" style:apply-style-name="P0"/></number:number-style>`;
const accounting = `${positive}<number:number-style style:name="NEG" style:volatile="true">${text("(")}${grouped(2)}${text(")")}</number:number-style>
  <number:number-style style:name="S">${text("-")}<style:map style:condition="value()&gt;0" style:apply-style-name="P0"/><style:map style:condition="value()&lt;0" style:apply-style-name="NEG"/></number:number-style>`;

test("style:map sections: sign handling and colour selection", () => {
  for (const [label, prefixes] of PREFIX_SETS) {
    const red = stylesOf(negativeRed, prefixes).get("S");
    assert.equal(red.info.code, '[>=0]#,##0.00;"-"#,##0.00', label);
    assert.deepEqual(formatOdfValue(red, n(-1234.5), NULL_DATE), { text: "-1,234.50", color: "#ff0000" }, `${label}: one minus sign, in red`);
    assert.deepEqual(formatOdfValue(red, n(1234.5), NULL_DATE), { text: "1,234.50" }, `${label}: positive, no colour`);
    assert.deepEqual(formatOdfValue(red, n(0), NULL_DATE), { text: "0.00" }, `${label}: zero takes the >=0 section`);
    const parentheses = stylesOf(accounting, prefixes).get("S");
    assert.equal(formatOdfValue(parentheses, n(-1234.5), NULL_DATE).text, "(1,234.50)", label);
    assert.equal(formatOdfValue(parentheses, n(1234.5), NULL_DATE).text, "1,234.50", label);
    assert.equal(formatOdfValue(parentheses, n(0), NULL_DATE).text, "-", `${label}: zero as a dash is still the number 0`);
  }
  // A section that positive values can also reach keeps the sign of negative values.
  const mixed = stylesOf(`${positive}<number:number-style style:name="S"><style:text-properties fo:color="#0000FF"/>${grouped(0)}<style:map style:condition="value()&gt;1000" style:apply-style-name="P0"/></number:number-style>`).get("S");
  assert.deepEqual(formatOdfValue(mixed, n(-5), NULL_DATE), { text: "-5", color: "#0000ff" });
  assert.deepEqual(formatOdfValue(mixed, n(5), NULL_DATE), { text: "5", color: "#0000ff" });
  assert.deepEqual(formatOdfValue(mixed, n(5000), NULL_DATE), { text: "5,000.00" });
});

test("constructs that cannot be reproduced exactly are unsupported", () => {
  const unsupported = [
    // formatValue does not read "# ?/16" as a fraction; a fixed denominator stays unsupported for now.
    ["fixed denominator", `<number:number-style style:name="S"><number:fraction number:min-integer-digits="0" number:min-numerator-digits="1" number:min-denominator-digits="2" number:denominator-value="16"/></number:number-style>`],
    ["embedded text", `<number:number-style style:name="S"><number:number number:decimal-places="2" number:min-integer-digits="1"><number:embedded-text number:position="3">-</number:embedded-text></number:number></number:number-style>`],
    ["fill character", `<number:number-style style:name="S">${num(2)}<number:fill-character> </number:fill-character></number:number-style>`],
    ["non-Gregorian calendar", `<number:date-style style:name="S"><number:day number:calendar="hijri"/>${text("/")}<number:month number:calendar="hijri"/></number:date-style>`],
    ["display factor 100", `<number:number-style style:name="S"><number:number number:decimal-places="0" number:min-integer-digits="1" number:display-factor="100"/></number:number-style>`],
    ["General (no decimal places)", `<number:number-style style:name="S"><number:number number:min-integer-digits="1"/></number:number-style>`],
    ["optional decimals", `<number:number-style style:name="S"><number:number number:decimal-places="2" number:min-decimal-places="0" number:min-integer-digits="1"/></number:number-style>`],
    ["optional decimals (loext)", `<number:number-style style:name="S"><number:number number:decimal-places="2" loext:min-decimal-places="1" number:min-integer-digits="1"/></number:number-style>`],
    ["decimal replacement", `<number:number-style style:name="S"><number:number number:decimal-places="2" number:min-integer-digits="1" number:decimal-replacement="--"/></number:number-style>`],
    ["era", `<number:date-style style:name="S"><number:era/>${text(" ")}<number:year/></number:date-style>`],
    ["quarter", `<number:date-style style:name="S"><number:quarter/>${text(" ")}<number:year/></number:date-style>`],
    ["week of year", `<number:date-style style:name="S"><number:week-of-year/></number:date-style>`],
    ["possessive month", `<number:date-style style:name="S"><number:day/>${text(" ")}<number:month number:textual="true" number:possessive-form="true"/></number:date-style>`],
    ["automatic order", `<number:date-style style:name="S" number:automatic-order="true"><number:day/>${text(".")}<number:month/></number:date-style>`],
    ["format source language", `<number:date-style style:name="S" number:format-source="language"><number:day/>${text(".")}<number:month/></number:date-style>`],
    ["transliteration", `<number:number-style style:name="S" number:transliteration-format="A">${num(0)}</number:number-style>`],
    ["unknown element", `<number:number-style style:name="S">${num(0)}<loext:fill-character> </loext:fill-character></number:number-style>`],
    ["empty currency symbol", `<number:currency-style style:name="S">${num(2)}<number:currency-symbol/></number:currency-style>`],
    ["unknown locale", `<number:number-style style:name="S" number:language="qq" number:country="QQ">${num(2)}</number:number-style>`],
    ["text style", `<number:text-style style:name="S"><number:text-content/></number:text-style>`],
    ["percentage without a percent sign", `<number:percentage-style style:name="S">${num(0)}</number:percentage-style>`],
    ["two number elements", `<number:number-style style:name="S">${num(0)}${text(" ")}${num(1)}</number:number-style>`],
    ["empty style", `<number:number-style style:name="S"/>`],
    ["minutes read as months", `<number:date-style style:name="S"><number:minutes number:style="long"/>${text(" ")}<number:day/></number:date-style>`],
    ["month read as minutes", `<number:date-style style:name="S"><number:hours/>${text(" ")}<number:month/></number:date-style>`],
    ["map target missing", `<number:number-style style:name="S">${num(0)}<style:map style:condition="value()&gt;=0" style:apply-style-name="NOPE"/></number:number-style>`],
    ["map condition not evaluable", `${positive}<number:number-style style:name="S">${num(0)}<style:map style:condition="is-true-formula(A1)" style:apply-style-name="P0"/></number:number-style>`],
    ["unsupported map target", `<number:number-style style:name="P" style:volatile="true"><number:number number:min-integer-digits="1"/></number:number-style><number:number-style style:name="S">${num(0)}<style:map style:condition="value()&gt;=0" style:apply-style-name="P"/></number:number-style>`],
  ];
  for (const [label, definition] of unsupported) {
    const format = stylesOf(definition).get("S");
    assert.ok(format, `${label}: read`);
    assert.equal(format.supported, false, `${label}: unsupported`);
    assert.equal(format.info.supported, false, `${label}: info`);
    assert.equal(format.info.code, undefined, `${label}: no code`);
    assert.equal(formatOdfValue(format, n(1.5), NULL_DATE)?.text, undefined, `${label}: no text`);
  }
  // A style whose sections cannot be chosen yields nothing, not the main section's colour.
  const unknownCondition = stylesOf(`${positive}<number:number-style style:name="S"><style:text-properties fo:color="#FF0000"/>${num(0)}<style:map style:condition="cell-content()&gt;0" style:apply-style-name="P0"/></number:number-style>`).get("S");
  assert.equal(formatOdfValue(unknownCondition, n(-1), NULL_DATE), undefined);
  // Values a supported style cannot show exactly: dates before 1900-03-01 and negative times.
  const iso = stylesOf(CASES.find(entry => entry[0] === "ISO date")[1]).get("S");
  assert.equal(formatOdfValue(iso, date("1900-02-28", 60), NULL_DATE).text, undefined);
  assert.equal(formatOdfValue(iso, date("1900-03-01", 61), NULL_DATE).text, "1900-03-01");
  assert.equal(formatOdfValue(iso, date("2026-10-28", 46323 - 1462), "1904-01-01").text, "2026-10-28", "serials relative to another null date");
  const clock = stylesOf(CASES.find(entry => entry[0] === "time of day")[1]).get("S");
  assert.equal(formatOdfValue(clock, time("-PT1H", -1 / 24), NULL_DATE).text, undefined);
  for (const value of [{ kind: "string", text: "x" }, { kind: "empty" }, { kind: "error", code: "#DIV/0!" }, { kind: "invalid", declaredType: "float" }, { kind: "date", lexical: "x" }]) assert.equal(formatOdfValue(iso, value, NULL_DATE), undefined, value.kind);
});

test("saved-text comparison unifies space and minus variants only", () => {
  assert.ok(sameDisplayText(`1${NBSP}234,50${NBSP}${EURO}`, `1 234,50 ${EURO}`));
  assert.ok(sameDisplayText(`1${NNBSP}234`, "1 234"));
  assert.ok(sameDisplayText(`${MINUS}5.00`, "-5.00"));
  assert.ok(sameDisplayText(" 3/4 ", "3/4"), "surrounding spaces (? placeholders) are trimmed");
  assert.ok(sameDisplayText("CHF 1'234.50", `CHF 1${String.fromCharCode(0x2019)}234.50`), "typographic and ASCII apostrophes (Swiss grouping)");
  assert.ok(!sameDisplayText("1,234.50", "1.234,50"));
  assert.ok(!sameDisplayText("12%", "12 %"));
});

// Parser integration: the M11 financial fixtures with the data styles a producer writes for them.
const FINANCIAL_STYLES = `
  <number:currency-style style:name="EUR"${lang("de", "DE")}>${grouped(2)}${text(" ")}<number:currency-symbol${lang("de", "DE")}>${EURO}</number:currency-symbol></number:currency-style>
  <number:currency-style style:name="JPY"${lang("ja", "JP")}><number:currency-symbol${lang("ja", "JP")}>${YEN}</number:currency-symbol>${grouped(0)}</number:currency-style>
  <number:currency-style style:name="CHF"${lang("de", "CH")}><number:currency-symbol${lang("de", "CH")}>CHF</number:currency-symbol>${text(" ")}${grouped(2)}</number:currency-style>
  <number:percentage-style style:name="PCT">${num(0)}${text("%")}</number:percentage-style>
  ${accounting.replace('style:name="S"', 'style:name="ACC"')}
  ${negativeRed.replace(/style:name="P0"/, 'style:name="P1"').replace('style:apply-style-name="P0"', 'style:apply-style-name="P1"').replace('style:name="S"', 'style:name="RED"')}
  <number:date-style style:name="DATE"><number:day/>${text(" ")}<number:month number:textual="true"/>${text(" ")}<number:year number:style="long"/></number:date-style>
  <number:time-style style:name="DUR" number:truncate-on-overflow="false"><number:hours/>${text(":")}<number:minutes number:style="long"/></number:time-style>
  <number:boolean-style style:name="BOOL"><number:boolean/></number:boolean-style>
  <number:number-style style:name="GEN"><number:number number:min-integer-digits="1"/></number:number-style>
  ${["EUR", "JPY", "CHF", "PCT", "ACC", "RED", "DATE", "DUR", "BOOL", "GEN"].map(name => `<style:style style:name="c${name}" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="${name}"/>`).join("")}`;
const cellXml = (style, attributes, saved) => `<table:table-cell table:style-name="c${style}" ${attributes}>${saved === undefined ? "" : `<text:p>${saved}</text:p>`}</table:table-cell>`;
const FINANCIAL_ROW = [
  ["EUR", 'office:value-type="currency" office:currency="EUR" office:value="1234.5"', `1.234,50 ${EURO}`],
  ["JPY", 'office:value-type="currency" office:currency="JPY" office:value="1235"', `${YEN}1,235`],
  ["CHF", 'office:value-type="currency" office:currency="CHF" office:value="1234.5"', "CHF 1'234.50"],
  ["PCT", 'office:value-type="percentage" office:value="0.15"', "15%"],
  ["ACC", 'office:value-type="float" office:value="-1234.5"', "(1,234.50)"],
  ["ACC", 'office:value-type="float" office:value="0"', "-"],
  ["RED", 'office:value-type="float" office:value="-1234.5"', "-1,234.50"],
  ["DATE", 'office:value-type="date" office:date-value="2026-10-28"', "28 Oct 2026"],
  ["DUR", 'office:value-type="time" office:time-value="PT36H30M00S"', "36:30"],
  ["BOOL", 'office:value-type="boolean" office:boolean-value="true"', "TRUE"],
];

function workbookOf(rows, prefixes = {}, styles = FINANCIAL_STYLES) {
  const workbook = parseOds(makeOds({ automatic: styles, body: `<table:table table:name="F">${rows.map(cells => `<table:table-row>${cells}</table:table-row>`).join("")}</table:table>` }, prefixes), "f.ods");
  return { workbook, cell: ref => workbook.sheets[0].cells.find(entry => entry.ref === ref), diagnostic: code => workbook.diagnostics.find(entry => entry.code === code) };
}

test("saved text is verified: computed text equals saved text for every supported financial fixture", () => {
  for (const [label, prefixes] of PREFIX_SETS) {
    const { workbook, cell, diagnostic } = workbookOf([FINANCIAL_ROW.map(([style, attributes, saved]) => cellXml(style, attributes, saved)).join("")], prefixes);
    FINANCIAL_ROW.forEach(([style, , saved], index) => {
      const entry = cell(`${String.fromCharCode(65 + index)}1`);
      assert.deepEqual(entry.data.display, { text: saved, source: "file" }, `${label} ${style}: saved text shown, no mismatch`);
      assert.equal(entry.display, saved, `${label} ${style}: legacy display`);
      assert.equal(workbook.styles[entry.style].numberFormatInfo.supported, true, `${label} ${style}: format supported`);
    });
    assert.equal(cell("G1").displayColor, "#ff0000", `${label}: negative red from the style:map section`);
    assert.equal(cell("E1").displayColor, undefined, label);
    assert.equal(diagnostic("odf.display.mismatch"), undefined, label);
    assert.equal(diagnostic("odf.display.raw"), undefined, label);
  }
});

test("cells without saved text are formatted (source computed); unsupported formats stay raw and are counted", () => {
  const cells = FINANCIAL_ROW.map(([style, attributes]) => cellXml(style, attributes)).join("") + cellXml("GEN", 'office:value-type="float" office:value="0.333333333333333"') + cellXml("DATE", 'office:value-type="date" office:date-value="1900-01-15"');
  const { cell, diagnostic } = workbookOf([cells, cellXml("GEN", 'office:value-type="float" office:value="2" table:number-columns-repeated="3"')]);
  FINANCIAL_ROW.forEach(([style, , saved], index) => {
    const entry = cell(`${String.fromCharCode(65 + index)}1`);
    assert.ok(sameDisplayText(entry.data.display.text, saved), `${style}: ${entry.data.display.text}`);
    assert.equal(entry.data.display.source, "computed", style);
    assert.equal(entry.display, entry.data.display.text, `${style}: legacy display`);
  });
  assert.ok(sameDisplayText(cell("C1").data.display.text, "CHF 1'234.50"), "Swiss grouping from the document locale (de-CH)");
  assert.deepEqual(cell("K1").data.display, { text: "0.333333333333333", source: "raw" }, "General: raw value");
  assert.deepEqual(cell("L1").data.display, { text: "1900-01-15", source: "raw" }, "date before 1900-03-01: raw value");
  assert.equal(diagnostic("odf.display.raw").count, 5, "the General and early-date cells plus three repeated copies");
});

test("a saved text that differs from the formatting is still shown and reported", () => {
  for (const [label, prefixes] of PREFIX_SETS) {
    const { cell, diagnostic } = workbookOf([[
      cellXml("EUR", 'office:value-type="currency" office:currency="EUR" office:value="1234.5"', `1.234,50${NBSP}${EURO}`),
      cellXml("EUR", 'office:value-type="currency" office:currency="EUR" office:value="1234.5" table:number-columns-repeated="2"', `1.234,49 ${EURO}`),
      cellXml("RED", 'office:value-type="float" office:value="-3"', `${MINUS}3.00`),
    ].join("")], prefixes);
    assert.deepEqual(cell("A1").data.display, { text: `1.234,50${NBSP}${EURO}`, source: "file" }, `${label}: NBSP is not a difference`);
    assert.deepEqual(cell("B1").data.display, { text: `1.234,49 ${EURO}`, source: "file", computed: `1.234,50 ${EURO}` }, `${label}: saved text wins`);
    assert.equal(cell("B1").display, `1.234,49 ${EURO}`, label);
    assert.deepEqual(cell("D1").data.display, { text: `${MINUS}3.00`, source: "file" }, `${label}: U+2212 is a minus sign`);
    assert.deepEqual(diagnostic("odf.display.mismatch"), { code: "odf.display.mismatch", severity: "accuracy", message: "2 cells: preview formatting differs from the saved text; the saved text is shown.", count: 2 }, label);
  }
});

test("data styles in content automatic styles take precedence over common styles of the same name", () => {
  const common = `<number:number-style style:name="N1">${num(3)}</number:number-style><style:style style:name="cN" style:family="table-cell" style:data-style-name="N1"/>`;
  const workbook = parseOds(makeOds({ styles: common, automatic: `<number:number-style style:name="N1">${num(1)}</number:number-style>`, body: `<table:table table:name="F"><table:table-row><table:table-cell table:style-name="cN" office:value-type="float" office:value="2"/></table:table-row></table:table>` }), "f.ods");
  assert.equal(workbook.sheets[0].cells[0].data.display.text, "2.0");
});

test("the document locale is used, never the viewer's (I-3)", () => {
  const styles = `<number:number-style style:name="N">${grouped(2)}</number:number-style><style:style style:name="cN" style:family="table-cell" style:data-style-name="N"/>`;
  const body = `<table:table table:name="F"><table:table-row><table:table-cell table:style-name="cN" office:value-type="float" office:value="1234.5"/></table:table-row></table:table>`;
  for (const [language, country, expected] of [["de", "DE", "1.234,50"], ["en", "US", "1,234.50"], ["fr", "CH", "1'234,50"]]) {
    const workbook = parseOds(makeOdfPackage({ content: spreadsheetContent({ body, automatic: styles }), styles: spreadsheetStyles({ language, country }) }), "f.ods");
    assert.equal(workbook.source.locale, `${language}-${country}`);
    const shown = workbook.sheets[0].cells[0].data.display;
    assert.equal(shown.source, "computed");
    assert.ok(sameDisplayText(shown.text, expected), `${language}-${country}: ${shown.text}`);
  }
});
