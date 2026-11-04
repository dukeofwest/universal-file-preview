import test from "node:test";
import assert from "node:assert/strict";
import { readdirSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { lengthToPoints, odfColor, parseBorder } from "../dist/parsers/odf/units.js";
import { parseXsdDouble, parseXsdDate, parseXsdDuration, dateSerial, readCellValue } from "../dist/parsers/odf/values.js";
import { savedCellText, cellRuns, cellHyperlink, annotationText } from "../dist/parsers/odf/text.js";
import { createStyleResolver } from "../dist/parsers/odf/styles.js";
import { ODF, EXT, FORMULA_NAMESPACES } from "../dist/parsers/odf/namespaces.js";
import { parseXmlNs, nsFirst, nsChildren } from "../dist/xml/ns-tree.js";
import { DiagnosticCollector } from "../dist/diagnostics.js";
import { odfXml } from "./helpers/odf.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
/** Every XML-based test runs with the canonical prefixes and with all prefixes renamed. */
const PREFIX_SETS = [
  ["canonical prefixes", {}],
  ["renamed prefixes", { office: "o", table: "t", text: "x", style: "s", fo: "f", svg: "g", number: "n", xlink: "l", calcext: "c", dc: "d" }],
];
const xml = (source, prefixes) => parseXmlNs(odfXml(source, prefixes));
const TABLEOOO = `xmlns:tableooo="${EXT.tableooo}"`;

// -----------------------------------------------------------------------------------------------
// Namespaces and units
// -----------------------------------------------------------------------------------------------

test("namespace constants cover the ODF core, formula languages and extensions", () => {
  assert.equal(ODF.table, "urn:oasis:names:tc:opendocument:xmlns:table:1.0");
  assert.equal(ODF.number, "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0");
  assert.equal(ODF.draw, "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0");
  assert.deepEqual([...FORMULA_NAMESPACES.values()], ["openformula", "openformula", "ooo", "msoxl"]);
  assert.equal(EXT.calcext, "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0");
});

test("lengthToPoints converts every ODF length unit and rejects everything else", () => {
  const cases = { "1in": 72, "2.54cm": 72, "25.4mm": 72, "12pt": 12, "1pc": 12, "16px": 12, ".5in": 36, "1e1pt": 10, "-1pt": -1, " 3pt ": 3, "0.06pt": 0.06 };
  for (const [value, expected] of Object.entries(cases)) assert.ok(Math.abs(lengthToPoints(value) - expected) < 1e-9, value);
  for (const value of [undefined, "10", "10%", "cm", "1 cm", "abc", "1.2.3pt", "NaNpt", "1em", ""]) assert.equal(lengthToPoints(value), undefined, String(value));
});

test("odfColor accepts #rrggbb only; parseBorder reads tokens in any order", () => {
  assert.equal(odfColor("#FF00aa"), "#ff00aa");
  for (const value of [undefined, "transparent", "#fff", "red", "#gg0000"]) assert.equal(odfColor(value), undefined, String(value));
  assert.deepEqual(parseBorder("0.06pt solid #000000"), { widthPt: 0.06, style: "solid", color: "#000000" });
  assert.deepEqual(parseBorder("solid #FF0000 1pt"), { widthPt: 1, style: "solid", color: "#ff0000" });
  assert.deepEqual(parseBorder("0.5pt dashed"), { widthPt: 0.5, style: "dashed", color: "#000000" });
  assert.deepEqual(parseBorder("thin double #00ff00"), { widthPt: 0.75, style: "double", color: "#00ff00" });
  for (const value of [undefined, "none", "hidden", "0.5pt none #000000", "1pt #000000", "1pt solid #000", "1pt wavy #000000"]) assert.equal(parseBorder(value), undefined, String(value));
});

// -----------------------------------------------------------------------------------------------
// Values
// -----------------------------------------------------------------------------------------------

test("parseXsdDouble accepts only finite xsd:double values", () => {
  const cases = { "1e3": 1000, ".5": 0.5, " 12.5 ": 12.5, "1.": 1, "+3": 3, "1234.5678": 1234.5678, "-1E-2": -0.01 };
  for (const [lexical, expected] of Object.entries(cases)) assert.equal(parseXsdDouble(lexical), expected, lexical);
  assert.ok(Object.is(parseXsdDouble("-0"), -0));
  for (const lexical of ["INF", "-INF", "NaN", "1,5", "", "0x10", "1e", "1 000", "1e400"]) assert.equal(parseXsdDouble(lexical), undefined, lexical);
});

test("parseXsdDate validates dates, times and time zones", () => {
  assert.deepEqual(parseXsdDate("2026-10-28"), { year: 2026, month: 10, day: 28, seconds: 0, hasTimezone: false });
  assert.deepEqual(parseXsdDate("2026-10-28T13:30:15.5"), { year: 2026, month: 10, day: 28, seconds: 48615.5, hasTimezone: false });
  assert.equal(parseXsdDate("2026-10-28T10:00:00+02:00").hasTimezone, true);
  assert.equal(parseXsdDate("2026-10-28T10:00:00Z").hasTimezone, true);
  assert.equal(parseXsdDate("2026-10-28T24:00:00").seconds, 86400);
  assert.equal(parseXsdDate("2024-02-29").day, 29);
  assert.equal(parseXsdDate("2000-02-29").day, 29);
  for (const lexical of ["2026-02-30", "2023-02-29", "1900-02-29", "2026-13-01", "2026-00-10", "2026-10-28T24:00:01", "2026-10-28T23:60:00", "26-10-28", "2026-10-28T10:00", "2026-10-28T10:00:00+15:00", "x"]) assert.equal(parseXsdDate(lexical), undefined, lexical);
});

test("parseXsdDuration measures fixed durations in days and flags calendar durations", () => {
  const close = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-12, `${actual} vs ${expected}`);
  close(parseXsdDuration("PT36H30M00S").days, 36.5 / 24);
  close(parseXsdDuration("PT12H").days, 0.5);
  close(parseXsdDuration("-PT12H").days, -0.5);
  close(parseXsdDuration("P2DT6H").days, 2.25);
  close(parseXsdDuration("PT0.5S").days, 0.5 / 86400);
  assert.equal(parseXsdDuration("PT0S").days, 0);
  assert.deepEqual(parseXsdDuration("P1Y"), { negative: false, hasYearOrMonth: true });
  assert.deepEqual(parseXsdDuration("-P1Y2MT3H"), { negative: true, hasYearOrMonth: true });
  for (const lexical of ["P", "PT", "P1H", "PT1Y", "12H", "", "PT1H2", "P-1D"]) assert.equal(parseXsdDuration(lexical), undefined, lexical);
});

test("dateSerial counts days from the null date without a phantom 29 February 1900", () => {
  const serial = (lexical, nullDate = "1899-12-30") => dateSerial(parseXsdDate(lexical), nullDate);
  assert.equal(serial("1899-12-30"), 0);
  assert.equal(serial("1900-01-01"), 2);
  // Consecutive days get consecutive serials (Excel's 1900 bug would give 59 for 28 February).
  assert.equal(serial("1900-02-28"), 60);
  assert.equal(serial("1900-03-01"), 61);
  assert.equal(serial("2026-10-28"), 46323, "the same serial the XLSX tests use");
  assert.equal(serial("2026-10-28T12:00:00"), 46323.5);
  assert.equal(serial("1904-01-01", "1904-01-01"), 0);
  assert.equal(serial("2026-10-28", "1904-01-01"), 46323 - 1462);
  assert.equal(serial("0050-01-02") - serial("0050-01-01"), 1, "years below 100 are not shifted to 19xx");
  assert.equal(serial("1899-12-29"), -1);
  assert.throws(() => serial("2026-10-28", "not-a-date"), RangeError);
});

/** Wraps cell attributes and content in a table:table-cell element. */
const cellXml = (attributes, content = "") => `<table:table-cell ${attributes}>${content}</table:table-cell>`;

test("readCellValue maps every ODF value type (spec 6.2)", () => {
  for (const [label, prefixes] of PREFIX_SETS) {
    const read = (attributes, savedText, nullDate = "1899-12-30") => {
      const collector = new DiagnosticCollector();
      const value = readCellValue(xml(cellXml(attributes), prefixes), savedText, collector, nullDate, { sheet: "S", cell: "A1" });
      return { value, codes: collector.list().map(entry => entry.code) };
    };
    const cases = [
      ['office:value-type="float" office:value="1234.5"', "999", { kind: "number", lexical: "1234.5", number: 1234.5 }, "number from the value, never the saved text"],
      ['office:value-type="float" office:value=" 12.50"', undefined, { kind: "number", lexical: " 12.50", number: 12.5 }, "lexical kept verbatim"],
      ['office:value-type="percentage" office:value="0.15"', "15%", { kind: "percentage", lexical: "0.15", number: 0.15 }],
      ['office:value-type="currency" office:currency="EUR" office:value="1234.5"', "1.234,50 EUR", { kind: "currency", lexical: "1234.5", number: 1234.5, currency: "EUR" }],
      ['office:value-type="currency" office:value="-7"', undefined, { kind: "currency", lexical: "-7", number: -7 }],
      ['office:value-type="date" office:date-value="2026-10-28"', "28.10.26", { kind: "date", lexical: "2026-10-28", serial: 46323 }],
      ['office:value-type="time" office:time-value="PT36H00M00S"', "36:00:00", { kind: "time", lexical: "PT36H00M00S", days: 1.5 }],
      ['office:value-type="time" office:time-value="P1M"', undefined, { kind: "time", lexical: "P1M" }],
      ['office:value-type="boolean" office:boolean-value="true"', "TRUE", { kind: "boolean", lexical: "true", boolean: true }],
      ['office:value-type="boolean" office:boolean-value="0"', undefined, { kind: "boolean", lexical: "0", boolean: false }],
      ['office:value-type="string" office:string-value="attribute"', "saved", { kind: "string", text: "attribute" }],
      ['office:value-type="string"', "saved", { kind: "string", text: "saved" }],
      ['office:value-type="string"', undefined, { kind: "string", text: "" }],
      ['table:style-name="ce1"', "1,000", { kind: "string", text: "1,000" }, "text that looks numeric stays a string"],
      ['table:style-name="ce1"', undefined, { kind: "empty" }],
      ['office:value-type="void"', undefined, { kind: "empty" }],
      ['office:value-type="float" office:value="0" calcext:value-type="error"', "#DIV/0!", { kind: "error", code: "#DIV/0!" }],
      ['office:value-type="string" calcext:value-type="error"', undefined, { kind: "error", code: "#ERROR" }],
    ];
    for (const [attributes, saved, expected, note] of cases) {
      const { value, codes } = read(attributes, saved);
      assert.deepEqual(value, expected, `${label}: ${note ?? attributes}`);
      assert.deepEqual(codes, [], `${label}: no diagnostics for ${attributes}`);
    }
    assert.ok(Object.is(read('office:value-type="float" office:value="-0"').value.number, -0));
    assert.deepEqual(read('office:value-type="float" office:value="0"').value, { kind: "number", lexical: "0", number: 0 }, `${label}: zero is not empty`);
    // A time zone keeps the lexical form and is reported.
    const zoned = read('office:value-type="date" office:date-value="2026-10-28T10:00:00+02:00"');
    assert.deepEqual([zoned.value.serial, zoned.codes], [46323 + 10 / 24, ["odf.value.timezone"]], label);
    // Without a valid null date, a date keeps its lexical form but has no serial.
    assert.deepEqual(read('office:value-type="date" office:date-value="2026-10-28"', undefined, "bad").value, { kind: "date", lexical: "2026-10-28" }, label);
  }
});

test("readCellValue reports every value that does not match its declared type (I-9)", () => {
  for (const [label, prefixes] of PREFIX_SETS) {
    const collector = new DiagnosticCollector();
    const invalid = [
      ['office:value-type="float"', { kind: "invalid", declaredType: "float" }],
      ['office:value-type="float" office:value="abc"', { kind: "invalid", declaredType: "float", lexical: "abc" }],
      ['office:value-type="percentage" office:value="INF"', { kind: "invalid", declaredType: "percentage", lexical: "INF" }],
      ['office:value-type="currency" office:value="1,5"', { kind: "invalid", declaredType: "currency", lexical: "1,5" }],
      ['office:value-type="date"', { kind: "invalid", declaredType: "date" }],
      ['office:value-type="date" office:date-value="2026-02-30"', { kind: "invalid", declaredType: "date", lexical: "2026-02-30" }],
      ['office:value-type="time" office:time-value="12h"', { kind: "invalid", declaredType: "time", lexical: "12h" }],
      ['office:value-type="boolean" office:boolean-value="yes"', { kind: "invalid", declaredType: "boolean", lexical: "yes" }],
      ['office:value-type="money" office:value="1"', { kind: "invalid", declaredType: "money" }],
    ];
    for (const [attributes, expected] of invalid) assert.deepEqual(readCellValue(xml(cellXml(attributes), prefixes), "shown", collector, "1899-12-30", { sheet: "S" }), expected, `${label}: ${attributes}`);
    assert.deepEqual(collector.list().map(entry => [entry.code, entry.severity, entry.count]), [["odf.value.invalid", "accuracy", invalid.length]], label);
  }
});

// -----------------------------------------------------------------------------------------------
// Text
// -----------------------------------------------------------------------------------------------

test("savedCellText follows ODF white-space rules", () => {
  for (const [label, prefixes] of PREFIX_SETS) {
    const text = content => savedCellText(xml(cellXml('office:value-type="string"', content), prefixes));
    assert.equal(text('<text:p>a<text:s text:c="3"/>b</text:p>'), "a   b", label);
    assert.equal(text("<text:p>a<text:tab/>b<text:line-break/>c</text:p>"), "a\tb\nc", label);
    assert.equal(text("<text:p>line 1</text:p><text:p>line 2</text:p>"), "line 1\nline 2", label);
    assert.equal(text('<text:p>x <text:span text:style-name="T1">bold <text:span text:style-name="T2">inner</text:span></text:span> y</text:p>'), "x bold inner y", label);
    assert.equal(text("<text:p>  a  \n\t b  </text:p>"), "a b", `${label}: literal white space collapses and is trimmed`);
    assert.equal(text("<text:p><text:s/>a<text:s/></text:p>"), " a ", `${label}: text:s keeps leading and trailing spaces`);
    assert.equal(text("<text:p>a <text:s/>b</text:p>"), "a  b", label);
    assert.equal(text('<text:p>a<text:s text:c="0"/>b<text:s text:c="99999"/>c</text:p>'), `a b${" ".repeat(10_000)}c`, `${label}: invalid counts are 1, huge counts are capped`);
    assert.equal(text('<office:annotation><dc:creator>Author</dc:creator><text:p>comment</text:p></office:annotation><text:p>value</text:p>'), "value", `${label}: comments are not display text`);
    assert.equal(text('<text:p>a<foo:bar xmlns:foo="urn:foo">b</foo:bar>c</text:p>'), "abc", `${label}: foreign element text is included`);
    assert.equal(text("<text:p/>"), "", label);
    assert.equal(text(""), undefined, `${label}: no text:p means no saved text`);
  }
});

test("cellRuns splits styled text and always matches the saved text", () => {
  const fonts = { T1: { bold: true }, T2: { italic: true, color: "#ff0000" } };
  const resolver = { textRunFont: name => fonts[name] };
  for (const [label, prefixes] of PREFIX_SETS) {
    const cell = content => xml(cellXml('office:value-type="string"', content), prefixes);
    const mixed = cell('<text:p>plain <text:span text:style-name="T1">bold</text:span> and <text:span text:style-name="T2">red</text:span></text:p><text:p>next</text:p>');
    const runs = cellRuns(mixed, resolver);
    assert.deepEqual(runs, [{ text: "plain " }, { text: "bold", font: { bold: true } }, { text: " and " }, { text: "red", font: { italic: true, color: "#ff0000" } }, { text: "\nnext" }], label);
    assert.equal(runs.map(run => run.text).join(""), savedCellText(mixed), label);
    assert.equal(cellRuns(cell("<text:p>no spans</text:p>"), resolver), undefined, label);
    assert.equal(cellRuns(cell('<text:p><text:span text:style-name="T1">all bold</text:span></text:p>'), resolver), undefined, `${label}: one style is not rich text`);
    assert.deepEqual(cellRuns(cell('<text:p>a<text:span text:style-name="Unknown">b</text:span></text:p>'), resolver), [{ text: "a" }, { text: "b" }], `${label}: unknown styles have no font`);
  }
});

test("cellHyperlink returns the first safe link; annotationText returns the comment text", () => {
  for (const [label, prefixes] of PREFIX_SETS) {
    const cell = content => xml(cellXml('office:value-type="string"', content), prefixes);
    assert.equal(cellHyperlink(cell('<text:p><text:a xlink:href="https://example.com/a">site</text:a></text:p>')), "https://example.com/a", label);
    assert.equal(cellHyperlink(cell('<text:p><text:a xlink:href="javascript:alert(1)">x</text:a><text:span><text:a xlink:href=" mailto:a@example.com ">m</text:a></text:span></text:p>')), "mailto:a@example.com", label);
    for (const target of ["javascript:alert(1)", "file:///etc/passwd", "data:text/html,x", "../other.ods", "#Sheet2.A1"]) assert.equal(cellHyperlink(cell(`<text:p><text:a xlink:href="${target}">x</text:a></text:p>`)), undefined, `${label}: ${target}`);
    assert.equal(cellHyperlink(cell("<text:p>no link</text:p>")), undefined, label);
    const commented = cell('<office:annotation><dc:creator>Author</dc:creator><dc:date>2026-10-28T10:00:00</dc:date><text:p>first</text:p><text:p>second <text:s/> line</text:p></office:annotation><text:p>value</text:p>');
    // Literal spaces collapse only with adjacent literal spaces, so both spaces around text:s are kept.
    assert.equal(annotationText(commented), "first\nsecond   line", label);
    assert.equal(annotationText(cell("<text:p>value</text:p>")), undefined, label);
  }
});

// -----------------------------------------------------------------------------------------------
// Styles
// -----------------------------------------------------------------------------------------------

const stylesDocument = `<office:document-styles>
  <office:font-face-decls>
    <style:font-face style:name="Liberation Sans" svg:font-family="'Liberation Sans'"/>
    <style:font-face style:name="DejaVu" svg:font-family="DejaVu Sans"/>
  </office:font-face-decls>
  <office:styles>
    <style:default-style style:family="table-cell"><style:text-properties style:font-name="Liberation Sans" fo:font-size="10pt" fo:language="de" fo:country="DE"/></style:default-style>
    <style:style style:name="Default" style:family="table-cell"/>
    <style:style style:name="Heading" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold" fo:font-size="16pt"/><style:table-cell-properties fo:border="0.06pt solid #000000" fo:background-color="#DDEEFF"/></style:style>
    <style:style style:name="Heading" style:family="paragraph"><style:text-properties fo:color="#FF0000"/></style:style>
    <style:style style:name="LoopA" style:family="table-cell" style:parent-style-name="LoopB"/>
    <style:style style:name="LoopB" style:family="table-cell" style:parent-style-name="LoopA"/>
    <number:currency-style style:name="N104"/>
  </office:styles>
</office:document-styles>`;
const automaticStyles = `<office:automatic-styles ${TABLEOOO}>
  <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Heading" style:data-style-name="N104">
    <style:table-cell-properties fo:border-left="2.49pt double #00FF00" style:vertical-align="middle" fo:wrap-option="wrap" style:rotation-angle="315" style:text-align-source="fix"/>
    <style:paragraph-properties fo:text-align="end"/>
    <style:text-properties fo:font-weight="normal" fo:font-style="italic" style:text-underline-style="solid" style:text-underline-type="double" style:text-line-through-style="solid" fo:color="#336699" style:text-position="super 58%"/>
  </style:style>
  <style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N2">
    <style:table-cell-properties style:text-align-source="value-type" style:diagonal-tl-br="0.74pt solid #123456" style:shrink-to-fit="true"/>
    <style:paragraph-properties fo:text-align="center"/>
  </style:style>
  <style:style style:name="Heading" style:family="table-cell"><style:text-properties fo:color="#00AA00"/></style:style>
  <style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="2.258cm"/></style:style>
  <style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.178in"/></style:style>
  <style:style style:name="ta1" style:family="table"><style:table-properties table:display="false" tableooo:tab-color="#FF0000"/></style:style>
  <style:style style:name="ta2" style:family="table"><style:table-properties table:display="true" table:tab-color="#0000FF"/></style:style>
  <style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" style:font-name="DejaVu"/></style:style>
  <number:number-style style:name="N2"/>
</office:automatic-styles>`;

function resolverFor(prefixes) {
  const styles = xml(stylesDocument, prefixes), automatic = xml(automaticStyles, prefixes), collector = new DiagnosticCollector();
  const fontFaces = nsChildren(styles, ODF.office, "font-face-decls");
  return { resolver: createStyleResolver(styles, automatic, fontFaces, collector), collector };
}

test("style resolution: default style, parent chain, per-attribute overrides and the spec 6.6 mapping", () => {
  for (const [label, prefixes] of PREFIX_SETS) {
    const { resolver, collector } = resolverFor(prefixes);
    assert.deepEqual(resolver.styles[0], { font: { family: "Liberation Sans", size: 10 } }, label);
    assert.deepEqual(resolver.defaultFont, { family: "Liberation Sans", size: 10 }, label);
    assert.equal(resolver.locale, "de-DE", label);
    assert.equal(resolver.cellStyleIndex(undefined), 0, label);
    assert.equal(resolver.cellStyleIndex("missing"), 0, label);
    const ce1 = resolver.cellStyleIndex("ce1");
    assert.equal(ce1, 1, label);
    assert.equal(resolver.cellStyleIndex("ce1"), 1, `${label}: cached`);
    assert.deepEqual(resolver.styles[ce1], {
      // family from the default style, size from Heading, bold overridden back to normal by ce1.
      font: { family: "Liberation Sans", size: 16, italic: true, underline: "double", strike: true, color: "#336699", vertAlign: "superscript" },
      fill: { type: "pattern", pattern: "solid", foreground: "#ddeeff" },
      borders: { top: { style: "thin", color: "#000000" }, bottom: { style: "thin", color: "#000000" }, left: { style: "double", color: "#00ff00" }, right: { style: "thin", color: "#000000" } },
      alignment: { horizontal: "right", vertical: "center", wrapText: true, textRotation: 135 },
      // M13: empty data styles cannot be reproduced; the effective locale is the document's.
      numberFormatInfo: { category: "currency", sourceName: "N104", supported: false, locale: "de-DE" },
    }, label);
    const ce2 = resolver.cellStyleIndex("ce2");
    assert.deepEqual(resolver.styles[ce2], {
      font: { family: "Liberation Sans", size: 10 },
      borders: { diagonal: { style: "thin", color: "#123456" }, diagonalDown: true },
      alignment: { shrinkToFit: true },
      numberFormatInfo: { category: "number", sourceName: "N2", supported: false, locale: "de-DE" },
    }, `${label}: value-type alignment means general, so no horizontal alignment`);
    assert.deepEqual(collector.list(), [], label);
  }
});

test("style names are scoped by family and automatic styles take precedence over common ones", () => {
  for (const [label, prefixes] of PREFIX_SETS) {
    const { resolver } = resolverFor(prefixes);
    // The cell reference "Heading" finds the automatic table-cell style, not the common one or the paragraph style.
    assert.deepEqual(resolver.styles[resolver.cellStyleIndex("Heading")].font, { family: "Liberation Sans", size: 10, color: "#00aa00" }, label);
    // ce1's parent "Heading" is the common table-cell style (bold, 16 pt), never the paragraph style's red.
    assert.notEqual(resolver.styles[resolver.cellStyleIndex("ce1")].font.color, "#ff0000", label);
  }
});

test("column, row, table, data style and text run lookups", () => {
  for (const [label, prefixes] of PREFIX_SETS) {
    const { resolver } = resolverFor(prefixes);
    assert.ok(Math.abs(resolver.columnWidthPoints("co1") - 2.258 * 72 / 2.54) < 1e-9, label);
    assert.equal(resolver.columnWidthPoints("missing"), undefined, label);
    assert.ok(Math.abs(resolver.rowProperties("ro1").heightPt - 0.178 * 72) < 1e-9, label);
    assert.deepEqual(resolver.rowProperties("missing"), {}, label);
    assert.deepEqual(resolver.tableProperties("ta1"), { hidden: true, tabColor: "#ff0000" }, label);
    assert.deepEqual(resolver.tableProperties("ta2"), { hidden: false, tabColor: "#0000ff" }, label);
    assert.deepEqual(resolver.tableProperties(undefined), { hidden: false }, label);
    assert.equal(resolver.dataStyleName("ce1"), "N104", label);
    assert.equal(resolver.dataStyleName("Default"), undefined, label);
    assert.deepEqual(resolver.textRunFont("T1"), { family: "DejaVu Sans", bold: true }, label);
    assert.equal(resolver.textRunFont("missing"), undefined, label);
  }
});

test("style inheritance cycles are cut and reported once", () => {
  for (const [label, prefixes] of PREFIX_SETS) {
    const { resolver, collector } = resolverFor(prefixes);
    const index = resolver.cellStyleIndex("LoopA");
    assert.deepEqual(resolver.styles[index], { font: { family: "Liberation Sans", size: 10 } }, label);
    resolver.cellStyleIndex("LoopA");
    assert.deepEqual(collector.list().map(entry => [entry.code, entry.severity, entry.count]), [["odf.style.cycle", "info", 1]], label);
  }
});

test("a document without a default cell style falls back to Liberation Sans 10 pt and says so", () => {
  const collector = new DiagnosticCollector(), resolver = createStyleResolver(undefined, undefined, [], collector);
  assert.deepEqual([resolver.styles, resolver.defaultFont, resolver.locale], [[{}], { family: "Liberation Sans", size: 10 }, undefined]);
  assert.deepEqual(collector.list().map(entry => entry.code), ["odf.style.default-missing"]);
  // A default-style alone (no "Default" style) is enough.
  const onlyDefault = createStyleResolver(xml('<office:document-styles><office:styles><style:default-style style:family="table-cell"><style:text-properties fo:font-family="Arial" fo:font-size="11pt" fo:language="zxx"/></style:default-style></office:styles></office:document-styles>', {}), undefined, [], collector);
  assert.deepEqual([onlyDefault.defaultFont, onlyDefault.locale], [{ family: "Arial", size: 11 }, undefined]);
  assert.equal(collector.list().length, 1);
});

test("text rotation and alignment edge cases", () => {
  const resolve = properties => {
    const automatic = xml(`<office:automatic-styles><style:style style:name="c" style:family="table-cell"><style:table-cell-properties ${properties.cell ?? ""}/><style:paragraph-properties ${properties.paragraph ?? ""}/></style:style></office:automatic-styles>`, {});
    const resolver = createStyleResolver(undefined, automatic, [], new DiagnosticCollector());
    return resolver.styles[resolver.cellStyleIndex("c")].alignment ?? {};
  };
  assert.equal(resolve({ cell: 'style:rotation-angle="90"' }).textRotation, 90);
  assert.equal(resolve({ cell: 'style:rotation-angle="270deg"' }).textRotation, 180);
  assert.equal(resolve({ cell: 'style:rotation-angle="360"' }).textRotation, undefined);
  assert.equal(resolve({ cell: 'style:rotation-angle="180"' }).textRotation, undefined, "upside-down text is not representable");
  assert.equal(resolve({ cell: 'style:rotation-angle="1.5rad"' }).textRotation, undefined);
  assert.equal(resolve({ paragraph: 'fo:text-align="start"' }).horizontal, "left", "no text-align-source means the fixed alignment applies");
  assert.equal(resolve({ paragraph: 'fo:text-align="justify"' }).horizontal, "justify");
  assert.equal(resolve({ cell: 'style:vertical-align="automatic"' }).vertical, undefined);
});

// -----------------------------------------------------------------------------------------------
// No prefixed-name comparisons
// -----------------------------------------------------------------------------------------------

test("compiled ODF modules contain no prefixed-name string literals", () => {
  const prefixes = ["office", "style", "text", "table", "fo", "number", "draw", "svg", "xlink", "meta", "calcext", "loext", "of", "manifest", "dc", "tableooo", "config", "chart"];
  const pattern = new RegExp(`["'\`](${prefixes.join("|")}):[A-Za-z]`);
  const files = [...readdirSync(join(root, "dist", "parsers", "odf")).filter(name => name.endsWith(".js")).map(name => join(root, "dist", "parsers", "odf", name)), join(root, "dist", "package", "odf-package.js")];
  assert.ok(files.length >= 6);
  for (const file of files) {
    const match = pattern.exec(readFileSync(file, "utf8"));
    assert.equal(match, null, `${file} contains ${match?.[0]}`);
  }
});
