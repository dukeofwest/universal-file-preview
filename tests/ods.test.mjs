import test from "node:test";
import assert from "node:assert/strict";
import { parseOds } from "../dist/parsers/odf/ods.js";
import { PreviewError } from "../dist/errors.js";
import { makeOds, makeOdfPackage, odfXml, spreadsheetContent, spreadsheetStyles, autoManifest, ODS_MEDIA_TYPE } from "./helpers/odf.mjs";

/** Every structural test runs with the canonical prefixes and with all prefixes renamed. */
const PREFIX_SETS = [
  ["canonical prefixes", {}],
  ["renamed prefixes", { office: "o", table: "t", text: "x", style: "s", fo: "f", svg: "g", number: "n", xlink: "l", calcext: "c", dc: "d", draw: "dr", meta: "m", of: "q", manifest: "mf" }],
];
const parse = (options, prefixes, parseOptions) => parseOds(makeOds(options, prefixes), "book.ods", parseOptions);
const cellAt = (sheet, ref) => sheet.cells.find(cell => cell.ref === ref);
const codes = workbook => workbook.diagnostics.map(entry => entry.code);
const diagnostic = (workbook, code) => workbook.diagnostics.find(entry => entry.code === code);
const fails = (run, code, message) => assert.throws(run, error => error instanceof PreviewError && error.code === code && message.test(error.message), String(message));

const float = (value, text = value, extra = "") => `<table:table-cell office:value-type="float" office:value="${value}" calcext:value-type="float"${extra}><text:p>${text}</text:p></table:table-cell>`;
const string = (text, extra = "") => `<table:table-cell office:value-type="string" calcext:value-type="string"${extra}><text:p>${text}</text:p></table:table-cell>`;
const table = (rows, attributes = 'table:name="Sheet1"', before = "") => `<table:table ${attributes}>${before}${rows}</table:table>`;
const row = (cells, attributes = "") => `<table:table-row ${attributes}>${cells}</table:table-row>`;

for (const [label, prefixes] of PREFIX_SETS) {
  test(`[${label}] a minimal one-sheet document`, () => {
    const workbook = parse({ body: table(row(string("Hello") + float("42")), 'table:name="Q1"', "<table:table-column table:number-columns-repeated=\"2\"/>") }, prefixes);
    assert.equal(workbook.kind, "workbook");
    assert.equal(workbook.title, "book.ods");
    assert.equal(workbook.sheets.length, 1);
    const [sheet] = workbook.sheets;
    assert.deepEqual([sheet.name, sheet.maxRow, sheet.maxColumn], ["Q1", 1, 2]);
    assert.deepEqual([cellAt(sheet, "A1").value, cellAt(sheet, "A1").type, cellAt(sheet, "A1").display], ["Hello", "str", "Hello"]);
    assert.deepEqual([cellAt(sheet, "B1").value, cellAt(sheet, "B1").type, cellAt(sheet, "B1").display], ["42", "n", "42"]);
    assert.deepEqual(workbook.source, { format: "ods", producer: "LibreOffice/24.8.4.2", formatVersion: "1.3", locale: "en-US" });
    assert.deepEqual(workbook.calculation, { nullDate: "1899-12-30" });
    assert.deepEqual(workbook.defaultFont, { family: "Liberation Sans", size: 10 });
    assert.deepEqual(codes(workbook), []);
  });

  test(`[${label}] template media type, multiple sheets and a hidden sheet`, () => {
    const options = { mimetype: "application/vnd.oasis.opendocument.spreadsheet-template", automatic: '<style:style style:name="ta1" style:family="table"><style:table-properties table:display="false"/></style:style><style:style style:name="ta2" style:family="table"><style:table-properties table:display="true"/></style:style>', body: table(row(string("a")), 'table:name="Visible" table:style-name="ta2"') + table(row(string("b")), 'table:name="Secret" table:style-name="ta1"') };
    const workbook = parse(options, prefixes);
    assert.equal(workbook.source.format, "ots");
    assert.deepEqual(workbook.sheets.map(sheet => [sheet.name, sheet.hidden ?? false]), [["Visible", false], ["Secret", true]]);
    assert.equal(diagnostic(workbook, "odf.hidden.sheets").count, 1);
  });

  test(`[${label}] LibreOffice-style trailing repeats stay compact and leave the bounds alone`, () => {
    const automatic = '<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="2.258cm"/></style:style><style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.178in"/></style:style><style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#FFFF00"/></style:style>';
    const body = table(
      row(string("Item") + float("5") + '<table:table-cell table:number-columns-repeated="1022"/>', 'table:style-name="ro1"')
      + row('<table:table-cell table:style-name="ce1" table:number-columns-repeated="1024"/>', 'table:style-name="ro1" table:number-rows-repeated="1048575"'),
      'table:name="Sheet1"', '<table:table-column table:style-name="co1" table:number-columns-repeated="1024" table:default-cell-style-name="Default"/>');
    const started = performance.now(), workbook = parse({ automatic, body }, prefixes), elapsed = performance.now() - started;
    const [sheet] = workbook.sheets;
    assert.deepEqual([sheet.maxRow, sheet.maxColumn], [1, 2], "styled empty rows do not extend the used range");
    assert.deepEqual(sheet.columns, [{ min: 1, max: 1024, widthPoints: 2.258 * 72 / 2.54 }]);
    assert.deepEqual(sheet.rowRanges, [{ min: 1, max: 1_048_576, height: 0.178 * 72 }], "adjacent identical row runs merge");
    const styled = sheet.cells.filter(cell => cell.value === "" && cell.data === undefined);
    // The yellow fill is materialised only inside the drawn grid: rows 2-60, columns A-Z.
    assert.equal(styled.length, 59 * 26);
    assert.ok(styled.every(cell => workbook.styles[cell.style].fill?.foreground === "#ffff00"));
    assert.equal(sheet.cells.length, 2 + 59 * 26);
    assert.ok(elapsed < 2000, `took ${elapsed.toFixed(0)} ms`);
  });

  test(`[${label}] repeated non-empty cells are materialised for every row and column`, () => {
    const workbook = parse({ body: table(row(float("7", "7", ' table:number-columns-repeated="3"'), 'table:number-rows-repeated="2"')) }, prefixes);
    assert.deepEqual(workbook.sheets[0].cells.map(cell => `${cell.ref}=${cell.value}`), ["A1=7", "B1=7", "C1=7", "A2=7", "B2=7", "C2=7"]);
  });

  test(`[${label}] the cell budget raises resource-limit`, () => {
    const options = { body: table(row(float("7", "7", ' table:number-columns-repeated="3"'), 'table:number-rows-repeated="2"')) };
    assert.equal(parse(options, prefixes, { limits: { maxCells: 6 } }).sheets[0].cells.length, 6);
    fails(() => parse(options, prefixes, { limits: { maxCells: 5 } }), "resource-limit", /^The spreadsheet exceeds the configured cell limit\.$/);
    fails(() => parse({ body: table(row(string("x".repeat(20)))) }, prefixes, { limits: { maxCellTextLength: 10 } }), "resource-limit", /cell's text exceeds/);
    assert.throws(() => parse(options, prefixes, { limits: { maxCells: 0 } }), RangeError);
    assert.throws(() => parse(options, prefixes, { limits: { unknown: 1 } }), RangeError);
  });

  test(`[${label}] merges and covered cells with content`, () => {
    const body = table(row(string("Merged", ' table:number-columns-spanned="2" table:number-rows-spanned="2"') + "<table:covered-table-cell/>") + row('<table:covered-table-cell/><table:covered-table-cell office:value-type="float" office:value="5"><text:p>5</text:p></table:covered-table-cell>'));
    const workbook = parse({ body }, prefixes), [sheet] = workbook.sheets;
    assert.deepEqual(sheet.merges, [{ top: 1, left: 1, bottom: 2, right: 2 }]);
    assert.equal(cellAt(sheet, "B2").data.covered, true);
    assert.equal(cellAt(sheet, "B2").value, "5");
    assert.equal(cellAt(sheet, "A2"), undefined);
    assert.equal(diagnostic(workbook, "odf.covered-cell-content").count, 1);
    assert.deepEqual([sheet.maxRow, sheet.maxColumn], [2, 2]);
  });

  test(`[${label}] column widths, row heights and hidden or filtered rows and columns`, () => {
    const automatic = '<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style><style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="2cm"/></style:style><style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="20pt"/></style:style>';
    const body = table(
      row(string("a")) + row(string("hidden"), 'table:visibility="collapse"') + row(string("filtered"), 'table:visibility="filter"') + row(string("b") + string("c") + string("d"), 'table:style-name="ro1"')
      + row("", 'table:number-rows-repeated="5"') + row("<table:table-cell/>", 'table:visibility="collapse" table:number-rows-repeated="3"'),
      'table:name="Sheet1"', '<table:table-column table:style-name="co1"/><table:table-column table:style-name="co2" table:visibility="collapse"/><table:table-column table:style-name="co2" table:visibility="filter"/>');
    const workbook = parse({ automatic, body }, prefixes), [sheet] = workbook.sheets;
    assert.deepEqual(sheet.columns, [{ min: 1, max: 1, widthPoints: 72 }, { min: 2, max: 2, widthPoints: 2 * 72 / 2.54, hidden: true }, { min: 3, max: 3, widthPoints: 2 * 72 / 2.54, hidden: true, filtered: true }]);
    assert.deepEqual(sheet.rowRanges, [{ min: 2, max: 2, hidden: true }, { min: 3, max: 3, hidden: true, filtered: true }, { min: 4, max: 4, height: 20 }, { min: 10, max: 12, hidden: true }]);
    assert.deepEqual([sheet.maxRow, sheet.maxColumn], [4, 3]);
    // Only hidden rows inside the used range (rows 1-4) are counted; rows 10-12 are outside it.
    assert.deepEqual([diagnostic(workbook, "odf.hidden.rows").count, diagnostic(workbook, "odf.hidden.filtered-rows").count, diagnostic(workbook, "odf.hidden.columns").count], [1, 1, 2]);
  });

  test(`[${label}] cell styles resolve through the cell, row and column defaults`, () => {
    const automatic = '<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold"/></style:style><style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-style="italic"/></style:style><style:style style:name="ce3" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:color="#FF0000"/></style:style>';
    const body = table(row(string("own", ' table:style-name="ce1"') + string("row default"), 'table:default-cell-style-name="ce2"') + row(string("plain") + string("column default")), 'table:name="Sheet1"', '<table:table-column/><table:table-column table:default-cell-style-name="ce3"/>');
    const workbook = parse({ automatic, body }, prefixes), [sheet] = workbook.sheets, font = ref => workbook.styles[cellAt(sheet, ref).style ?? 0].font;
    assert.equal(font("A1").bold, true);
    assert.equal(font("B1").italic, true, "row default-cell-style-name applies before the column's");
    assert.equal(font("A2").bold, undefined);
    assert.equal(font("B2").color, "#ff0000");
    assert.equal(cellAt(sheet, "A2").style, undefined, "the default style index is omitted");
  });

  test(`[${label}] hyperlinks, comments and comment-only cells`, () => {
    const body = table(row(
      '<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://example.com/report">Report</text:a></text:p></table:table-cell>'
      + '<table:table-cell office:value-type="string"><text:p><text:a xlink:href="javascript:alert(1)">bad</text:a></text:p></table:table-cell>'
      + '<table:table-cell office:value-type="float" office:value="1"><office:annotation><dc:creator>Auditor</dc:creator><text:p>Check this</text:p></office:annotation><text:p>1</text:p></table:table-cell>'
      + '<table:table-cell><office:annotation><text:p>Note only</text:p></office:annotation></table:table-cell>'));
    const [sheet] = parse({ body }, prefixes).sheets;
    assert.deepEqual(sheet.hyperlinks, { A1: "https://example.com/report" });
    assert.deepEqual(sheet.comments, { C1: "Check this", D1: "Note only" });
    assert.equal(cellAt(sheet, "D1"), undefined, "a comment alone is not cell content");
    assert.equal(cellAt(sheet, "C1").display, "1", "the comment is not part of the displayed text");
  });

  test(`[${label}] workbook and sheet named ranges`, () => {
    const workbookNames = '<table:named-expressions><table:named-range table:name="Total" table:base-cell-address="$Sheet1.$A$1" table:cell-range-address="$Sheet1.$B$2"/><table:named-expression table:name="Rate" table:base-cell-address="$Sheet1.$A$1" table:expression="of:=0.19"/></table:named-expressions>';
    const sheetNames = '<table:named-expressions><table:named-range table:name="Local" table:cell-range-address="$Sheet1.$A$1:.$A$3"/></table:named-expressions>';
    const workbook = parse({ body: table(row(string("x")) + sheetNames) + workbookNames }, prefixes);
    assert.deepEqual(workbook.names, [{ name: "Local", sheet: "Sheet1", ref: "$Sheet1.$A$1:.$A$3" }, { name: "Total", ref: "$Sheet1.$B$2" }, { name: "Rate", expression: odfXml('<r a="of:=0.19"/>', prefixes).match(/a="([^"]*)"/)[1] }]);
  });

  test(`[${label}] formula dialects are resolved by namespace`, () => {
    const formulas = ['of:=SUM([.A1])', 'oooc:=SUM([.A1])', 'msoxl:=SUM(A1)', 'zz:=SUM(A1)', '=1+1'];
    const cells = formulas.map(text => `<table:table-cell table:formula="${text}" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>`).join("");
    const body = table(row(cells), 'table:name="Sheet1" xmlns:oooc="http://openoffice.org/2004/formula" xmlns:msoxl="http://schemas.microsoft.com/office/excel/formula"');
    const workbook = parse({ body }, prefixes), [sheet] = workbook.sheets;
    const formula = ref => cellAt(sheet, ref).data.formula;
    assert.deepEqual([formula("A1").dialect, formula("A1").expression, formula("A1").namespace], ["openformula", "=SUM([.A1])", "urn:oasis:names:tc:opendocument:xmlns:of:1.2"]);
    assert.deepEqual([formula("B1").dialect, formula("B1").expression], ["ooo", "=SUM([.A1])"]);
    assert.deepEqual([formula("C1").dialect, formula("C1").expression], ["msoxl", "=SUM(A1)"]);
    assert.deepEqual([formula("D1").dialect, formula("D1").expression], ["unknown", "zz:=SUM(A1)"]);
    assert.deepEqual([formula("E1").dialect, formula("E1").expression], ["unknown", "=1+1"]);
    assert.equal(cellAt(sheet, "A1").formula, "=SUM([.A1])");
    assert.equal(diagnostic(workbook, "odf.formula.unknown-dialect").count, 2);
    assert.equal(diagnostic(workbook, "odf.formula.cached").count, 5);
  });

  test(`[${label}] drawings, charts and conditional formatting are reported, not rendered`, () => {
    const shapes = '<table:shapes><draw:frame draw:name="Picture"><draw:image xlink:href="Pictures/a.png"/></draw:frame><draw:frame draw:name="Chart"><draw:object xlink:href="./Object 1"/></draw:frame></table:shapes>';
    const conditional = '<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="Sheet1.A1:Sheet1.A2"><calcext:condition calcext:apply-style-name="Bad" calcext:value="&lt;0"/></calcext:conditional-format></calcext:conditional-formats>';
    const manifest = autoManifest(ODS_MEDIA_TYPE, ["content.xml", "styles.xml", "meta.xml", "Pictures/a.png", "Object 1/content.xml"]).replace("</manifest:manifest>", '<manifest:file-entry manifest:full-path="Object 1/" manifest:media-type="application/vnd.oasis.opendocument.chart"/></manifest:manifest>');
    const workbook = parse({ body: table(shapes + row(string("x")) + conditional), extra: { "Pictures/a.png": "png", "Object 1/content.xml": "<x/>" }, manifest }, prefixes);
    assert.equal(diagnostic(workbook, "odf.unsupported.drawings").count, 2);
    assert.equal(diagnostic(workbook, "odf.unsupported.charts").count, 1);
    assert.ok(diagnostic(workbook, "odf.unsupported.conditional-formatting"));
    assert.deepEqual(workbook.sheets[0].cells.map(cell => cell.ref), ["A1"]);
    // A style:map on a used cell style also counts as conditional formatting.
    const mapped = parse({ automatic: '<style:style style:name="ce1" style:family="table-cell"><style:map style:condition="cell-content()&lt;0" style:apply-style-name="Bad" style:base-cell-address="Sheet1.A1"/></style:style>', body: table(row(float("1", "1", ' table:style-name="ce1"'))) }, prefixes);
    assert.ok(diagnostic(mapped, "odf.unsupported.conditional-formatting"));
  });

  test(`[${label}] external links and pivot tables are reported`, () => {
    const workbook = parse({ body: table(row(string("x")), 'table:name="Sheet1"', '<table:table-source xlink:href="file:///other.ods" table:table-name="Data"/>') + '<table:data-pilot-tables><table:data-pilot-table table:name="P"/></table:data-pilot-tables>' }, prefixes);
    assert.ok(diagnostic(workbook, "odf.external-links"));
    assert.ok(diagnostic(workbook, "odf.unsupported.pivot"));
  });

  test(`[${label}] content beyond the sheet limits is dropped and reported`, () => {
    const body = table(row(string("first")) + row("", 'table:number-rows-repeated="1048575"') + row(string("too far")) + row(`<table:table-cell table:number-columns-repeated="16384"/>${string("too wide")}`));
    const workbook = parse({ body }, prefixes), [sheet] = workbook.sheets;
    assert.deepEqual(sheet.cells.map(cell => cell.ref), ["A1"]);
    assert.ok(diagnostic(workbook, "odf.bounds.clipped"));
  });
}

test("package and content errors are reported with clear codes", () => {
  const styles = spreadsheetStyles();
  const encrypted = autoManifest(ODS_MEDIA_TYPE, ["styles.xml"]).replace("</manifest:manifest>", '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum="x"/></manifest:file-entry></manifest:manifest>');
  fails(() => parseOds(makeOdfPackage({ content: "\u0001 encrypted bytes", styles, manifest: encrypted }), "x.ods"), "unsupported-feature", /^Password-protected/);
  fails(() => parseOds(makeOdfPackage({ mimetype: "application/vnd.oasis.opendocument.text", content: spreadsheetContent({ body: "" }), styles }), "x.odt"), "unsupported-format", /^Not an OpenDocument spreadsheet\.$/);
  fails(() => parseOds(makeOdfPackage({ content: odfXml("<office:document-content><a></b></office:document-content>"), styles }), "x.ods"), "parse-error", /^content\.xml: Mismatched closing tag b/);
  fails(() => parseOds(makeOdfPackage({ content: spreadsheetContent({ body: "" }), styles: "<broken" }), "x.ods"), "parse-error", /^styles\.xml: /);
  fails(() => parseOds(makeOdfPackage({ content: odfXml("<office:document-styles/>"), styles }), "x.ods"), "invalid-file", /^content\.xml is not an OpenDocument content stream\.$/);
  fails(() => parseOds(makeOdfPackage({ content: spreadsheetContent({ body: "" }), styles }), "x.ods"), "invalid-file", /^ODS workbook contains no readable sheets\.$/);
  fails(() => parseOds(makeOdfPackage({ styles }), "x.ods"), "invalid-file", /no content\.xml/);
  const controller = new AbortController(); controller.abort();
  fails(() => parseOds(makeOds({ body: table(row(string("x"))) }), "x.ods", { signal: controller.signal }), "aborted", /cancelled/);
  fails(() => parseOds(makeOds({ body: table(row(string("a"))) + table(row(string("b"))) }), "x.ods", { limits: { maxSheets: 1 } }), "resource-limit", /more sheets/);
});

test("a 100,000-row by 10-column spreadsheet parses within the specification's budget", () => {
  const cells = Array.from({ length: 10 }, (_, column) => `<table:table-cell office:value-type="float" office:value="${column + 0.5}" calcext:value-type="float"><text:p>${column + 0.5}</text:p></table:table-cell>`).join("");
  const rowXml = `<table:table-row>${cells}</table:table-row>`;
  const bytes = makeOds({ body: `<table:table table:name="Big"><table:table-column table:number-columns-repeated="10"/>${rowXml.repeat(100_000)}</table:table>` });
  const before = process.memoryUsage().heapUsed, started = performance.now();
  const workbook = parseOds(bytes, "big.ods");
  const elapsed = performance.now() - started, growth = process.memoryUsage().heapUsed - before;
  assert.equal(workbook.sheets[0].cells.length, 1_000_000);
  assert.equal(workbook.sheets[0].maxRow, 100_000);
  console.log(`# parseOds: ${(bytes.length / 1048576).toFixed(1)} MiB package, 1,000,000 cells in ${elapsed.toFixed(0)} ms; heap growth ${(growth / 1048576).toFixed(0)} MiB`);
  assert.ok(elapsed < 15_000, `took ${elapsed.toFixed(0)} ms`);
  assert.ok(growth < 600 * 1048576, `heap grew ${(growth / 1048576).toFixed(0)} MiB`);
});
