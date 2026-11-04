// Flat OpenDocument spreadsheets (.fods; implementation specification M15).
import test from "node:test";
import assert from "node:assert/strict";
import { parseFods } from "../dist/parsers/odf/flat.js";
import { parseOds } from "../dist/parsers/odf/ods.js";
import { PreviewError } from "../dist/errors.js";
import { streamXml } from "../dist/xml/xml-stream.js";
import { createDefaultRendererRegistry } from "../dist/renderers/registry.js";
import { OdsRenderer } from "../dist/renderers/ods-renderer.js";
import { detectFormat } from "../dist/detection/detect-format.js";
import { flatOdsXml, makeFods, makeOdfPackage, makeOds, odfXml, spreadsheetContent, spreadsheetStyles } from "./helpers/odf.mjs";

const PREFIX_SETS = [["canonical prefixes", {}], ["renamed prefixes", { office: "o", table: "t", text: "x", style: "s", fo: "f", svg: "g", number: "n", xlink: "l", calcext: "c", dc: "d", draw: "dr", meta: "m", of: "q", config: "cfg", loext: "lo" }]];
const EURO = String.fromCharCode(0x20ac);
const fails = (run, code, message) => assert.throws(run, error => error instanceof PreviewError && error.code === code && message.test(error.message), String(message));

const cell = (attributes, text) => `<table:table-cell ${attributes}>${text === undefined ? "" : `<text:p>${text}</text:p>`}</table:table-cell>`;
const row = (cells, attributes = "") => `<table:table-row${attributes ? ` ${attributes}` : ""}>${cells}</table:table-row>`;
const table = (rows, attributes = 'table:name="Sheet1"', before = "") => `<table:table ${attributes}>${before}${rows}</table:table>`;

/** Workbooks exercising the M11 core behaviour; each is parsed as a package and as a flat file. */
const FIXTURES = {
  values: {
    body: table(row([
      cell('office:value-type="float" office:value="1234.5" calcext:value-type="float"', "1,234.50"),
      cell('office:value-type="percentage" office:value="0.15"', "15%"),
      cell('office:value-type="currency" office:currency="EUR" office:value="-7"', `-7.00 ${EURO}`),
      cell('office:value-type="date" office:date-value="2026-10-28T13:05:00"', "28.10.26 13:05"),
      cell('office:value-type="time" office:time-value="PT36H30M00S"'),
      cell('office:value-type="boolean" office:boolean-value="false"', "FALSE"),
      cell('office:value-type="string"', "text"),
      cell('office:value-type="float" office:value="0" calcext:value-type="error"', "#DIV/0!"),
      cell('office:value-type="float" office:value="abc"', "?"),
      cell('office:value-type="float" office:value="0"'),
    ].join(""))),
  },
  layout: {
    automatic: '<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="3cm"/></style:style><style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.8cm"/></style:style><style:style style:name="taH" style:family="table"><style:table-properties table:display="false"/></style:style><style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:background-color="#ffcc00"/></style:style>',
    body: table(
      row(cell('office:value-type="string" table:number-columns-spanned="2" table:number-rows-spanned="2"', "Merged") + "<table:covered-table-cell/>" + cell('office:value-type="float" office:value="7" table:number-columns-repeated="3"', "7"))
      + row('<table:covered-table-cell/><table:covered-table-cell office:value-type="float" office:value="5"><text:p>5</text:p></table:covered-table-cell>', 'table:style-name="ro1"')
      + row(cell('office:value-type="string"', "hidden"), 'table:visibility="collapse"') + row(cell('office:value-type="string"', "filtered"), 'table:visibility="filter"')
      + row(cell('table:style-name="ce1" table:number-columns-repeated="4"'), 'table:number-rows-repeated="3"'),
      'table:name="Layout"', '<table:table-column table:style-name="co1"/><table:table-column table:style-name="co1" table:visibility="collapse"/><table:table-column table:number-columns-repeated="3" table:default-cell-style-name="ce1"/>')
      + table(row(cell('office:value-type="string"', "secret")), 'table:name="Hidden" table:style-name="taH"'),
  },
  styles: {
    styles: '<style:style style:name="Heading" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold" fo:font-size="14pt"/></style:style><number:number-style style:name="N4"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:number-style>',
    automatic: '<number:number-style style:name="N5P0" style:volatile="true"><number:number number:decimal-places="1" number:min-decimal-places="1" number:min-integer-digits="1"/></number:number-style><number:number-style style:name="N5"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text><number:number number:decimal-places="1" number:min-decimal-places="1" number:min-integer-digits="1"/><style:map style:condition="value()&gt;=0" style:apply-style-name="N5P0"/></number:number-style><style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Heading" style:data-style-name="N4"><style:table-cell-properties fo:border="0.06pt solid #000000"/></style:style><style:style style:name="ce3" style:family="table-cell" style:data-style-name="N5"/><style:style style:name="T1" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>',
    fontFaces: '<style:font-face style:name="DejaVu Sans" svg:font-family="\'DejaVu Sans\'"/>',
    body: table(row([
      cell('table:style-name="ce2" office:value-type="float" office:value="1234.5"', "1,234.50"),
      cell('table:style-name="ce2" office:value-type="float" office:value="2"'),
      cell('table:style-name="ce3" office:value-type="float" office:value="-3"', "-3.0"),
      cell('table:style-name="ce3" office:value-type="float" office:value="4"', "4.1"),
      '<table:table-cell office:value-type="string"><text:p>plain <text:span text:style-name="T1">italic</text:span></text:p></table:table-cell>',
    ].join(""))),
  },
  formulas: {
    body: '<table:calculation-settings table:precision-as-shown="true"><table:null-date table:date-value="1904-01-01"/></table:calculation-settings>'
      + '<table:named-expressions><table:named-range table:name="Total" table:base-cell-address="$Sheet1.$A$1" table:cell-range-address="$Sheet1.$B$2"/><table:named-expression table:name="Rate" table:base-cell-address="$Sheet1.$A$1" table:expression="of:=0.19"/></table:named-expressions>'
      + table(row([
        cell('table:formula="of:=SUM([.B1:.B3])" office:value-type="float" office:value="6"', "6"),
        cell('table:formula="of:=A1*2"'),
        cell('table:formula="of:=TRANSPOSE([.A1:.A2])" table:number-matrix-columns-spanned="1" table:number-matrix-rows-spanned="2" office:value-type="float" office:value="1"', "1"),
        cell('office:value-type="date" office:date-value="2026-10-28"', "2026-10-28"),
      ].join(""))),
  },
  text: {
    body: table(row([
      '<table:table-cell office:value-type="string"><text:p><text:a xlink:href="https://example.com/report">Report</text:a></text:p></table:table-cell>',
      '<table:table-cell office:value-type="float" office:value="1"><office:annotation><dc:creator>Auditor</dc:creator><text:p>Check this</text:p></office:annotation><text:p>1</text:p></table:table-cell>',
      '<table:table-cell office:value-type="string"><text:p>line one</text:p><text:p>line two</text:p></table:table-cell>',
      '<table:table-cell office:value-type="string"><text:p>a<text:s text:c="3"/>b<text:tab/>c</text:p></table:table-cell>',
    ].join(""))),
  },
  unsupported: {
    automatic: '<style:style style:name="ceC" style:family="table-cell"><style:map style:condition="cell-content()&gt;5" style:apply-style-name="Default"/></style:style>',
    body: table(row(cell('table:style-name="ceC" office:value-type="float" office:value="9"', "9") + '<table:table-cell><draw:frame draw:name="pic"><draw:image xlink:href="Pictures/x.png"/></draw:frame></table:table-cell>'), 'table:name="Sheet1"', '<table:table-source xlink:href="file:///other.ods" table:table-name="Data"/>')
      + '<table:dde-links><table:dde-link/></table:dde-links><table:data-pilot-tables><table:data-pilot-table table:name="P"/></table:data-pilot-tables>',
  },
};
const SETTINGS = '<config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry><config:config-item-map-named config:name="Tables"><config:config-item-map-entry config:name="Layout"><config:config-item config:name="VerticalSplitMode" config:type="short">2</config:config-item><config:config-item config:name="VerticalSplitPosition" config:type="int">1</config:config-item><config:config-item config:name="ZoomValue" config:type="int">80</config:config-item></config:config-item-map-entry></config:config-item-map-named><config:config-item config:name="ActiveTable" config:type="string">Hidden</config:config-item></config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set>';

/** The package equivalent of flatOdsXml, with settings.xml when `settings` is given. */
function packageOf(options, prefixes) {
  const meta = odfXml(`<office:document-meta office:version="1.3"><office:meta><meta:generator>LibreOffice/24.8.4.2$Windows_X86_64 LibreOffice_project/test</meta:generator></office:meta></office:document-meta>`, prefixes);
  const settings = options.settings === undefined ? undefined : odfXml(`<office:document-settings office:version="1.3"><office:settings>${options.settings}</office:settings></office:document-settings>`, prefixes);
  return makeOdfPackage({ content: spreadsheetContent(options, prefixes), styles: spreadsheetStyles(options, prefixes), meta, ...(settings !== undefined ? { settings } : {}) });
}

for (const [label, prefixes] of PREFIX_SETS) {
  test(`[${label}] a flat file yields exactly the workbook of the equivalent package`, () => {
    for (const [name, fixture] of Object.entries({ ...FIXTURES, settings: { ...FIXTURES.layout, settings: SETTINGS } })) {
      const zipped = parseOds(packageOf(fixture, prefixes), "book");
      const flat = parseFods(makeFods(fixture, prefixes), "book");
      assert.deepEqual(flat, { ...zipped, source: { ...zipped.source, format: "fods" } }, `${label} ${name}`);
      assert.ok(flat.sheets[0].cells.length, `${label} ${name}: cells were read`);
    }
  });
}

test("the flat workbook carries the M11 values, styles, view settings and diagnostics", () => {
  const values = parseFods(makeFods(FIXTURES.values), "v.fods");
  const at = ref => values.sheets[0].cells.find(entry => entry.ref === ref);
  assert.deepEqual(values.source, { format: "fods", producer: "LibreOffice/24.8.4.2", formatVersion: "1.3", locale: "en-US" });
  assert.deepEqual(at("A1").data.value, { kind: "number", lexical: "1234.5", number: 1234.5 });
  assert.deepEqual(at("A1").data.display, { text: "1,234.50", source: "file" });
  assert.deepEqual(at("H1").data.value, { kind: "error", code: "#DIV/0!" });
  assert.deepEqual(at("J1").data.value, { kind: "number", lexical: "0", number: 0 }, "zero is not empty");
  const styled = parseFods(makeFods(FIXTURES.styles), "s.fods");
  const s = ref => styled.sheets[0].cells.find(entry => entry.ref === ref);
  assert.deepEqual(s("B1").data.display, { text: "2.00", source: "computed" }, "common data style from office:styles");
  assert.equal(s("C1").displayColor, "#ff0000", "automatic data style with a map");
  assert.deepEqual(s("D1").data.display, { text: "4.1", source: "file", computed: "4.0" });
  assert.equal(styled.styles[s("A1").style].font.bold, true, "parent style from office:styles");
  const viewed = parseFods(makeFods({ ...FIXTURES.layout, settings: SETTINGS }), "l.fods");
  assert.equal(viewed.activeSheet, 1);
  assert.deepEqual(viewed.sheets[0].frozen, { rows: 1, columns: 0 });
  assert.equal(viewed.sheets[0].zoom, 0.8);
  const codes = parseFods(makeFods(FIXTURES.unsupported), "u.fods").diagnostics.map(entry => entry.code);
  for (const code of ["odf.unsupported.conditional-formatting", "odf.unsupported.drawings", "odf.external-links", "odf.unsupported.pivot"]) assert.ok(codes.includes(code), code);
});

test("the root must be office:document with a spreadsheet media type", () => {
  fails(() => parseFods(makeFods({ body: table(row("")), mimetype: "application/vnd.oasis.opendocument.text" }), "x.fods"), "unsupported-format", /^Not an OpenDocument spreadsheet\.$/);
  fails(() => parseFods(new TextEncoder().encode(flatOdsXml({ body: table(row("")) }).replace(/ office:mimetype="[^"]*"/, "")), "x.fods"), "unsupported-format", /^Not an OpenDocument spreadsheet\.$/);
  fails(() => parseFods(new TextEncoder().encode(spreadsheetContent({ body: table(row("")) })), "x.fods"), "invalid-file", /^The file is not a flat OpenDocument document\.$/);
  fails(() => parseFods(new TextEncoder().encode('<office:document xmlns:office="urn:example:not-odf" office:mimetype="application/vnd.oasis.opendocument.spreadsheet"/>'), "x.fods"), "invalid-file", /^The file is not a flat OpenDocument document\.$/);
  fails(() => parseFods(makeFods({ body: "" }), "x.fods"), "invalid-file", /no readable sheets/);
  fails(() => parseFods(new TextEncoder().encode("<office:document"), "x.fods"), "parse-error", /^flat spreadsheet: /);
  const template = parseFods(makeFods({ body: table(row(cell('office:value-type="string"', "t"))), mimetype: "application/vnd.oasis.opendocument.spreadsheet-template" }), "t.fods");
  assert.equal(template.source.format, "fods", "a flat template is a flat spreadsheet");
});

test("decoding follows the package text rules; the size limit applies to the whole file", () => {
  const xml = flatOdsXml({ body: table(row(cell('office:value-type="string"', `caf${String.fromCharCode(0xe9)}`))) });
  const utf16 = new Uint8Array(2 + xml.length * 2);
  utf16[0] = 0xff; utf16[1] = 0xfe;
  for (let i = 0; i < xml.length; i++) { utf16[2 + i * 2] = xml.charCodeAt(i) & 0xff; utf16[3 + i * 2] = xml.charCodeAt(i) >> 8; }
  assert.equal(parseFods(utf16, "u.fods").sheets[0].cells[0].value, `caf${String.fromCharCode(0xe9)}`, "UTF-16LE by byte-order mark");
  const bom = new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode(xml)]);
  assert.equal(parseFods(bom, "b.fods").sheets[0].cells[0].value, `caf${String.fromCharCode(0xe9)}`, "UTF-8 byte-order mark");
  fails(() => parseFods(new TextEncoder().encode(xml.replace('encoding="UTF-8"', 'encoding="ISO-8859-1"')), "l.fods"), "unsupported-feature", /^The flat spreadsheet uses the unsupported encoding ISO-8859-1\.$/);
  const broken = new TextEncoder().encode(xml); broken[broken.length - 20] = 0xff;
  fails(() => parseFods(broken, "i.fods"), "invalid-file", /^The flat spreadsheet is not valid UTF-8 text\.$/);
  const bytes = new TextEncoder().encode(xml);
  fails(() => parseFods(bytes, "big.fods", { limits: { maxContentBytes: bytes.length - 1 } }), "resource-limit", /^The flat spreadsheet exceeds the configured size limit\.$/);
  assert.ok(parseFods(bytes, "ok.fods", { limits: { maxContentBytes: bytes.length } }));
  assert.throws(() => parseFods(bytes, "x.fods", { limits: { maxStuff: 1 } }), RangeError);
});

test("embedded images and objects are not kept, and inline objects never end the sheet", () => {
  const base64 = "iVBORw0KGgo".repeat(20_000);
  const chart = `<office:document office:mimetype="application/vnd.oasis.opendocument.chart"><office:automatic-styles><style:style style:name="ce1" style:family="table-cell"><style:text-properties fo:color="#00ff00"/></style:style></office:automatic-styles><office:body><office:chart><table:table table:name="local-table"><table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row></table:table></office:chart></office:body></office:document>`;
  const body = table(
    `<table:shapes><draw:frame draw:name="logo"><draw:image><office:binary-data>${base64}</office:binary-data></draw:image></draw:frame><draw:frame draw:name="chart"><draw:object>${chart}</draw:object></draw:frame></table:shapes>`
    + row(cell('office:value-type="string"', "before") + `<table:table-cell office:value-type="float" office:value="2"><draw:frame draw:name="c2"><draw:object>${chart}</draw:object></draw:frame><draw:frame draw:name="p"><draw:image><office:binary-data>${base64}</office:binary-data></draw:image></draw:frame><text:p>2</text:p></table:table-cell>`)
    + row(cell('office:value-type="string"', "after")));
  const automatic = '<style:style style:name="ce1" style:family="table-cell"><style:text-properties fo:color="#123456"/></style:style>';
  // A cell capture limit of 50 elements: the embedded chart document inside a cell is not kept.
  const workbook = parseFods(makeFods({ body: body + table(row(cell('office:value-type="string"', "second")), 'table:name="Second"'), automatic }), "e.fods", { limits: { maxCaptureElements: 50 } });
  assert.deepEqual(workbook.sheets.map(sheet => sheet.name), ["Sheet1", "Second"], "the inline chart's table is not a sheet and does not end Sheet1");
  assert.deepEqual(workbook.sheets[0].cells.map(entry => [entry.ref, entry.value]), [["A1", "before"], ["B1", "2"], ["A2", "after"]]);
  assert.equal(workbook.styles.some(style => style.font?.color === "#00ff00"), false, "the embedded document's automatic styles are not the spreadsheet's");
  const codes = workbook.diagnostics.filter(entry => entry.code.startsWith("odf.unsupported")).map(entry => [entry.code, entry.count]);
  assert.deepEqual(codes, [["odf.unsupported.drawings", 4], ["odf.unsupported.charts", 2]]);
});

test("streamXml keeps elements but drops their content when emptyInCapture matches", () => {
  const captured = [];
  const handler = { start: element => element.name.local === "cell" ? "capture" : undefined, captured: node => captured.push(node) };
  streamXml('<r><cell a="1"><bin type="png">AAAA<x>BB</x>CC</bin><keep>text</keep><bin/></cell></r>', handler, { emptyInCapture: (uri, local) => local === "bin" });
  const [node] = captured;
  assert.deepEqual(node.children.map(child => [child.local, child.attributes.map(attribute => attribute.value), child.children]), [["bin", ["png"], []], ["keep", [], ["text"]], ["bin", [], []]]);
  assert.throws(() => streamXml("<r><cell><bin><x></bin></cell></r>", handler, { emptyInCapture: () => true }), PreviewError, "content is still checked");
});

test("fods files are routed to the ODS renderer and detected from their content", () => {
  const bytes = makeFods(FIXTURES.values);
  assert.equal(detectFormat("book.fods", "", bytes), "fods");
  assert.ok(createDefaultRendererRegistry().get("fods") instanceof OdsRenderer);
  assert.equal(detectFormat("book.ods", "", makeOds({ body: table(row("")) })), "ods", "packages are unaffected");
});
