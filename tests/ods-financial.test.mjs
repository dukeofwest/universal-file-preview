// Financial accuracy suite for the ODS parser (implementation specification §6, milestone M11).
// Every case asserts the typed value, the displayed text and its source, the formula, the legacy
// fields the grid renderer reads, and the diagnostics, each separately.
import test from "node:test";
import assert from "node:assert/strict";
import { parseOds } from "../dist/parsers/odf/ods.js";
import { makeOds } from "./helpers/odf.mjs";

const EURO = String.fromCharCode(0x20ac), YEN = String.fromCharCode(0xa5);
const OF = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";
const PREFIX_SETS = [["canonical prefixes", {}], ["renamed prefixes", { office: "o", table: "t", text: "x", calcext: "c", of: "q", style: "s", fo: "f", svg: "g", meta: "m", manifest: "mf" }]];

/** Parses one sheet whose first row holds `cells` (and optional extra rows and spreadsheet-level elements). */
function sheetOf(cells, { prefixes = {}, rows = "", before = "" } = {}) {
  const workbook = parseOds(makeOds({ body: `${before}<table:table table:name="F"><table:table-row>${cells}</table:table-row>${rows}</table:table>` }, prefixes), "f.ods");
  return { workbook, sheet: workbook.sheets[0], cell: ref => workbook.sheets[0].cells.find(candidate => candidate.ref === ref), codes: () => workbook.diagnostics.map(entry => entry.code) };
}

/** Asserts the five aspects of one cell separately. */
function expectCell(cell, { value, display, formula, legacy }, label) {
  assert.ok(cell, `${label}: cell exists`);
  assert.deepEqual(cell.data.value, value, `${label}: typed value`);
  assert.deepEqual(cell.data.display, display, `${label}: display text and source`);
  assert.deepEqual(cell.data.formula, formula, `${label}: formula`);
  const actual = { value: cell.value, ...(cell.type !== undefined ? { type: cell.type } : {}), ...(cell.display !== undefined ? { display: cell.display } : {}), ...(cell.formula !== undefined ? { formula: cell.formula } : {}) };
  assert.deepEqual(actual, legacy, `${label}: legacy fields read by the renderer`);
}

for (const [label, prefixes] of PREFIX_SETS) {
  test(`[${label}] currencies keep their ISO code, their stored value and the saved text (I-1, I-2, I-10)`, () => {
    const cells = [
      `<table:table-cell office:value-type="currency" office:currency="EUR" office:value="1234.5" calcext:value-type="currency"><text:p>1.234,50 ${EURO}</text:p></table:table-cell>`,
      `<table:table-cell office:value-type="currency" office:currency="JPY" office:value="1235" calcext:value-type="currency"><text:p>${YEN}1,235</text:p></table:table-cell>`,
      `<table:table-cell office:value-type="currency" office:currency="CHF" office:value="1234.5"><text:p>CHF 1'234.50</text:p></table:table-cell>`,
      // A style showing "$" for a EUR value: the code is preserved, the saved text is shown as saved.
      `<table:table-cell office:value-type="currency" office:currency="EUR" office:value="1234.5"><text:p>$1,234.50</text:p></table:table-cell>`,
      // The saved text never feeds the number.
      `<table:table-cell office:value-type="currency" office:currency="USD" office:value="1234.5"><text:p>999</text:p></table:table-cell>`,
    ].join("");
    const { cell, codes } = sheetOf(cells, { prefixes });
    expectCell(cell("A1"), { value: { kind: "currency", lexical: "1234.5", number: 1234.5, currency: "EUR" }, display: { text: `1.234,50 ${EURO}`, source: "file" }, legacy: { value: "1234.5", type: "n", display: `1.234,50 ${EURO}` } }, "EUR");
    expectCell(cell("B1"), { value: { kind: "currency", lexical: "1235", number: 1235, currency: "JPY" }, display: { text: `${YEN}1,235`, source: "file" }, legacy: { value: "1235", type: "n", display: `${YEN}1,235` } }, "JPY");
    expectCell(cell("C1"), { value: { kind: "currency", lexical: "1234.5", number: 1234.5, currency: "CHF" }, display: { text: "CHF 1'234.50", source: "file" }, legacy: { value: "1234.5", type: "n", display: "CHF 1'234.50" } }, "CHF");
    expectCell(cell("D1"), { value: { kind: "currency", lexical: "1234.5", number: 1234.5, currency: "EUR" }, display: { text: "$1,234.50", source: "file" }, legacy: { value: "1234.5", type: "n", display: "$1,234.50" } }, "symbol differs from code");
    expectCell(cell("E1"), { value: { kind: "currency", lexical: "1234.5", number: 1234.5, currency: "USD" }, display: { text: "999", source: "file" }, legacy: { value: "1234.5", type: "n", display: "999" } }, "text never feeds the number");
    assert.deepEqual(codes(), []);
  });

  test(`[${label}] percentages, negatives, zero shown as a dash, and text that looks numeric (I-8)`, () => {
    const cells = [
      '<table:table-cell office:value-type="percentage" office:value="0.15"><text:p>15%</text:p></table:table-cell>',
      '<table:table-cell office:value-type="float" office:value="-1234.5"><text:p>(1,234.50)</text:p></table:table-cell>',
      '<table:table-cell office:value-type="float" office:value="0"><text:p>-</text:p></table:table-cell>',
      '<table:table-cell office:value-type="string"><text:p>1,000</text:p></table:table-cell>',
      '<table:table-cell/>',
      '<table:table-cell office:value-type="float" office:value="0"><text:p>0</text:p></table:table-cell>',
    ].join("");
    const { cell, codes } = sheetOf(cells, { prefixes });
    expectCell(cell("A1"), { value: { kind: "percentage", lexical: "0.15", number: 0.15 }, display: { text: "15%", source: "file" }, legacy: { value: "0.15", type: "n", display: "15%" } }, "percentage is stored as a fraction");
    expectCell(cell("B1"), { value: { kind: "number", lexical: "-1234.5", number: -1234.5 }, display: { text: "(1,234.50)", source: "file" }, legacy: { value: "-1234.5", type: "n", display: "(1,234.50)" } }, "negative in parentheses");
    expectCell(cell("C1"), { value: { kind: "number", lexical: "0", number: 0 }, display: { text: "-", source: "file" }, legacy: { value: "0", type: "n", display: "-" } }, "zero displayed as a dash is still zero");
    expectCell(cell("D1"), { value: { kind: "string", text: "1,000" }, display: { text: "1,000", source: "file" }, legacy: { value: "1,000", type: "str", display: "1,000" } }, "numeric-looking text stays text");
    assert.equal(cell("E1"), undefined, "an empty cell is not materialised, so it is not zero");
    expectCell(cell("F1"), { value: { kind: "number", lexical: "0", number: 0 }, display: { text: "0", source: "file" }, legacy: { value: "0", type: "n", display: "0" } }, "zero is a value");
    assert.deepEqual(codes(), []);
  });

  test(`[${label}] dates, date-times and durations use ISO values and the document's null date`, () => {
    const cells = [
      '<table:table-cell office:value-type="date" office:date-value="2026-10-28"><text:p>28.10.26</text:p></table:table-cell>',
      '<table:table-cell office:value-type="date" office:date-value="2026-10-28T13:30:15"><text:p>28.10.2026 13:30:15</text:p></table:table-cell>',
      '<table:table-cell office:value-type="time" office:time-value="PT36H00M00S"><text:p>36:00:00</text:p></table:table-cell>',
    ].join("");
    const standard = sheetOf(cells, { prefixes });
    expectCell(standard.cell("A1"), { value: { kind: "date", lexical: "2026-10-28", serial: 46323 }, display: { text: "28.10.26", source: "file" }, legacy: { value: "46323", type: "n", display: "28.10.26" } }, "date");
    const serial = 46323 + 48615 / 86400;
    expectCell(standard.cell("B1"), { value: { kind: "date", lexical: "2026-10-28T13:30:15", serial }, display: { text: "28.10.2026 13:30:15", source: "file" }, legacy: { value: String(serial), type: "n", display: "28.10.2026 13:30:15" } }, "date-time with seconds");
    expectCell(standard.cell("C1"), { value: { kind: "time", lexical: "PT36H00M00S", days: 1.5 }, display: { text: "36:00:00", source: "file" }, legacy: { value: "1.5", type: "n", display: "36:00:00" } }, "duration over 24 hours");
    assert.equal(standard.workbook.calculation.nullDate, "1899-12-30");
    // A 1904 null date shifts serials, not the ISO values.
    const shifted = sheetOf(cells, { prefixes, before: '<table:calculation-settings><table:null-date table:date-value="1904-01-01"/></table:calculation-settings>' });
    assert.equal(shifted.cell("A1").data.value.serial, 46323 - 1462);
    assert.equal(shifted.cell("A1").data.value.lexical, "2026-10-28");
    assert.equal(shifted.workbook.calculation.nullDate, "1904-01-01");
  });

  test(`[${label}] invalid values are marked and reported, never coerced (I-9)`, () => {
    const { cell, workbook } = sheetOf('<table:table-cell office:value-type="float" office:value="abc"><text:p>abc</text:p></table:table-cell><table:table-cell office:value-type="float"><text:p>12</text:p></table:table-cell>', { prefixes });
    expectCell(cell("A1"), { value: { kind: "invalid", declaredType: "float", lexical: "abc" }, display: { text: "abc", source: "file" }, legacy: { value: "abc", type: "str", display: "abc" } }, "invalid float lexical");
    expectCell(cell("B1"), { value: { kind: "invalid", declaredType: "float" }, display: { text: "12", source: "file" }, legacy: { value: "12", type: "str", display: "12" } }, "missing value attribute; the saved text is shown as text, not as a number");
    assert.deepEqual(workbook.diagnostics.find(entry => entry.code === "odf.value.invalid"), { code: "odf.value.invalid", severity: "accuracy", message: "2 cells contain values that do not match their declared type.", count: 2 });
  });

  test(`[${label}] cached formula results are shown as saved; missing ones are never invented (I-5, I-13)`, () => {
    const cells = [
      '<table:table-cell table:formula="of:=SUM([.B2:.B3])" office:value-type="float" office:value="3" calcext:value-type="float"><text:p>3</text:p></table:table-cell>',
      '<table:table-cell table:formula="of:=[.A1]*2"/>',
      '<table:table-cell table:formula="of:=1/0" office:value-type="float" office:value="0" calcext:value-type="error"><text:p>#DIV/0!</text:p></table:table-cell>',
    ].join("");
    const { cell, workbook } = sheetOf(cells, { prefixes });
    const cached = { text: `${prefixes.of ?? "of"}:=SUM([.B2:.B3])`, expression: "=SUM([.B2:.B3])", dialect: "openformula", namespace: OF, result: "cached" };
    expectCell(cell("A1"), { value: { kind: "number", lexical: "3", number: 3 }, display: { text: "3", source: "file" }, formula: cached, legacy: { value: "3", type: "n", display: "3", formula: "=SUM([.B2:.B3])" } }, "cached result");
    // No saved value or text: the formula is not calculated, and the legacy display is omitted so the grid shows "Cannot calculate".
    const missing = { text: `${prefixes.of ?? "of"}:=[.A1]*2`, expression: "=[.A1]*2", dialect: "openformula", namespace: OF, result: "missing" };
    expectCell(cell("B1"), { value: { kind: "empty" }, display: { text: "", source: "raw" }, formula: missing, legacy: { value: "", formula: "=[.A1]*2" } }, "missing result");
    const error = { text: `${prefixes.of ?? "of"}:=1/0`, expression: "=1/0", dialect: "openformula", namespace: OF, result: "cached" };
    expectCell(cell("C1"), { value: { kind: "error", code: "#DIV/0!" }, display: { text: "#DIV/0!", source: "file" }, formula: error, legacy: { value: "#DIV/0!", type: "e", display: "#DIV/0!", formula: "=1/0" } }, "error result");
    const diagnostics = Object.fromEntries(workbook.diagnostics.map(entry => [entry.code, entry]));
    assert.deepEqual(diagnostics["odf.formula.cached"], { code: "odf.formula.cached", severity: "accuracy", message: "Formula results are shown as last saved by LibreOffice/24.8.4.2; they were not recalculated.", count: 3 });
    assert.deepEqual(diagnostics["odf.formula.missing-result"], { code: "odf.formula.missing-result", severity: "accuracy", message: '1 formula cell has no saved result and is shown as "Cannot calculate".', count: 1 });
  });

  test(`[${label}] precision as shown is reported for the whole workbook`, () => {
    const { workbook } = sheetOf('<table:table-cell office:value-type="float" office:value="1.005"><text:p>1.01</text:p></table:table-cell>', { prefixes, before: '<table:calculation-settings table:precision-as-shown="true"/>' });
    assert.equal(workbook.calculation.precisionAsShown, true);
    assert.equal(workbook.diagnostics.find(entry => entry.code === "odf.precision-as-shown").severity, "accuracy");
    assert.equal(workbook.sheets[0].cells[0].data.value.number, 1.005, "the full stored value is kept");
  });

  test(`[${label}] hidden rows are counted only inside the used range (I-7)`, () => {
    const rows = '<table:table-row table:visibility="collapse"><table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell></table:table-row>'
      + '<table:table-row><table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell></table:table-row>'
      + '<table:table-row table:number-rows-repeated="6"><table:table-cell/></table:table-row>'
      + '<table:table-row table:visibility="collapse" table:number-rows-repeated="3"><table:table-cell/></table:table-row>';
    const { workbook, sheet, cell } = sheetOf('<table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>', { prefixes, rows });
    assert.equal(sheet.maxRow, 3);
    assert.deepEqual(sheet.rowRanges, [{ min: 2, max: 2, hidden: true }, { min: 10, max: 12, hidden: true }], "hidden data stays in the model");
    assert.equal(cell("A2").data.value.number, 2, "hidden values are kept");
    assert.deepEqual(workbook.diagnostics.find(entry => entry.code === "odf.hidden.rows"), { code: "odf.hidden.rows", severity: "accuracy", message: "1 hidden row in the used range.", count: 1 });
  });

  test(`[${label}] a number without saved text is shown raw and reported, never formatted with the viewer's locale (I-3)`, () => {
    const { cell, workbook } = sheetOf('<table:table-cell office:value-type="float" office:value="1234.5678"/><table:table-cell office:value-type="date" office:date-value="2026-10-28"/><table:table-cell office:value-type="boolean" office:boolean-value="true"/>', { prefixes });
    expectCell(cell("A1"), { value: { kind: "number", lexical: "1234.5678", number: 1234.5678 }, display: { text: "1234.5678", source: "raw" }, legacy: { value: "1234.5678", type: "n", display: "1234.5678" } }, "raw number");
    expectCell(cell("B1"), { value: { kind: "date", lexical: "2026-10-28", serial: 46323 }, display: { text: "2026-10-28", source: "raw" }, legacy: { value: "46323", type: "n", display: "2026-10-28" } }, "raw date");
    expectCell(cell("C1"), { value: { kind: "boolean", lexical: "true", boolean: true }, display: { text: "TRUE", source: "raw" }, legacy: { value: "1", type: "b", display: "TRUE" } }, "raw boolean");
    assert.deepEqual(workbook.diagnostics.find(entry => entry.code === "odf.display.raw"), { code: "odf.display.raw", severity: "accuracy", message: "2 values had no saved display text and are shown unformatted.", count: 2 });
  });
}
