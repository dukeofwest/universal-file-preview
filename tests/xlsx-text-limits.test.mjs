import test from "node:test";
import assert from "node:assert/strict";
import { parseXlsx } from "../dist/parsers/xlsx.js";
import { evaluateFormula } from "../dist/parsers/xlsx-formula.js";
import { formatValue } from "../dist/parsers/xlsx-number-format.js";
import { makeZip } from "./helpers/zip.mjs";

// Text limits for XLSX (0.2.2): one cell may hold at most 32,767 characters (Excel's limit), and the text of all
// cells together, counted for every cell that uses it, at most 64 Mi UTF-16 code units. Shared strings and shared
// number formats let a small file expand into far more text than it contains; 0.2.1 rendered 142 million characters
// from a 195 KiB file.

const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const col = index => String.fromCharCode(65 + Math.floor(index / 26)) + String.fromCharCode(65 + index % 26);
const failsWith = (code, pattern) => error => error?.name === "PreviewError" && error.code === code && pattern.test(error.message);
/** A one-sheet workbook; `cells` are full <c> elements laid out 100 per row. */
function workbook(cells, { shared, styles } = {}) {
  const rows = [];
  for (let start = 0; start < cells.length; start += 100) rows.push(`<row r="${start / 100 + 1}">${cells.slice(start, start + 100).map((cell, index) => cell(`${col(index)}${start / 100 + 1}`)).join("")}</row>`);
  return new Uint8Array(makeZip({
    "xl/workbook.xml": `<workbook ${S}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/>${shared ? '<Relationship Id="ss" Type="x/sharedStrings" Target="sharedStrings.xml"/>' : ""}${styles ? '<Relationship Id="st" Type="x/styles" Target="styles.xml"/>' : ""}</Relationships>`,
    ...(shared ? { "xl/sharedStrings.xml": `<sst ${S}>${shared.map(text => `<si><t>${text}</t></si>`).join("")}</sst>` } : {}),
    ...(styles ? { "xl/styles.xml": styles } : {}),
    "xl/worksheets/sheet1.xml": `<worksheet ${S}><sheetData>${rows.join("")}</sheetData></worksheet>`,
  }, { deflate: true }));
}
const sharedCell = index => ref => `<c r="${ref}" t="s"><v>${index}</v></c>`;
const repeat = (count, cell) => Array.from({ length: count }, () => cell);
const parse = bytes => parseXlsx(bytes, "x.xlsx");

test("one cell may hold exactly 32,767 characters, from a shared string, an inline string or a formula's cached text", () => {
  const max = "a".repeat(32_767), over = "a".repeat(32_768);
  assert.equal(parse(workbook([sharedCell(0)], { shared: [max] })).sheets[0].cells[0].value.length, 32_767);
  assert.throws(() => parse(workbook([sharedCell(0)], { shared: [over] })), failsWith("resource-limit", /cell text is longer than 32,767/));
  const inline = text => workbook([ref => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`]);
  assert.equal(parse(inline(max)).sheets[0].cells[0].value.length, 32_767);
  assert.throws(() => parse(inline(over)), failsWith("resource-limit", /cell text is longer/));
  const cached = text => workbook([ref => `<c r="${ref}" t="str"><f>"x"</f><v>${text}</v></c>`]);
  assert.equal(parse(cached(max)).sheets[0].cells[0].value.length, 32_767);
  assert.throws(() => parse(cached(over)), failsWith("resource-limit", /cell text is longer/));
});

test("the workbook budget counts a shared string again for every cell that uses it", () => {
  // Each reference is within the per-cell limit; only the aggregate exceeds the budget (64 Mi / 32,767 = 2,048.06).
  const text = "b".repeat(32_767);
  for (const [cells, accepted] of [[1, true], [1000, true], [2048, true], [2049, false], [5000, false]]) {
    const bytes = workbook(repeat(cells, sharedCell(0)), { shared: [text] }), started = performance.now();
    if (accepted) assert.equal(parse(bytes).sheets[0].cells.length, cells, `${cells} references`);
    else assert.throws(() => parse(bytes), failsWith("resource-limit", /workbook text exceeds 67,108,864/), `${cells} references`);
    assert.ok(performance.now() - started < 3000, `${cells} references decided quickly`);
  }
});

test("the 0.2.1 attack file (100,000 cells sharing one 1,000,000-character string) is refused promptly", () => {
  const bytes = workbook(repeat(100_000, sharedCell(0)), { shared: ["y".repeat(1_000_000)] }), started = performance.now();
  assert.ok(bytes.length < 256 * 1024, `the file is small (${bytes.length} bytes)`);
  assert.throws(() => parse(bytes), failsWith("resource-limit", /cell text is longer than 32,767/));
  assert.ok(performance.now() - started < 3000);
});

test("number formats cannot multiply text: repeated @ placeholders and long literals are limited", () => {
  assert.equal(formatValue("x".repeat(10_000), "@@@").text.length, 30_000);
  assert.throws(() => formatValue("x".repeat(10_000), "@@@@"), failsWith("resource-limit", /Formatted cell text exceeds 32,767/));
  // A text cell formatted with "@" ten times: each cell would display 10 x 10,000 characters.
  const styles = `<styleSheet ${S}><numFmts count="1"><numFmt numFmtId="164" formatCode="${"@".repeat(10)}"/></numFmts><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>`;
  assert.throws(() => parse(workbook([ref => `<c r="${ref}" s="1" t="s"><v>0</v></c>`], { shared: ["z".repeat(10_000)], styles })), failsWith("resource-limit", /32,767/));
  // A 30,000-character literal applied to many numeric cells: each display is within the cell limit, the total is not.
  const literal = `<styleSheet ${S}><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;${"q".repeat(30_000)}&quot;0"/></numFmts><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>`;
  const started = performance.now();
  assert.throws(() => parse(workbook(repeat(5000, ref => `<c r="${ref}" s="1"><v>1</v></c>`), { styles: literal })), failsWith("resource-limit", /workbook text exceeds/));
  assert.ok(performance.now() - started < 3000);
});

test("formula text results longer than 32,767 characters are #VALUE!, checked before the text is built", () => {
  const long = "w".repeat(20_000);
  const context = { sheet: "S", row: 1, column: 1, cell: () => long, name: () => undefined };
  const value = formula => { const result = evaluateFormula(formula, context); return typeof result === "string" ? `text(${result.length})` : result?.code ?? result; };
  assert.equal(value('A1&"x"'), "text(20001)");
  assert.equal(value("A1&A1"), "#VALUE!");
  assert.equal(value("REPT(\"ab\",16383)"), "text(32766)");
  assert.equal(value("REPT(\"ab\",16384)"), "#VALUE!");
  assert.equal(value("REPT(A1,2)"), "#VALUE!");
  assert.equal(value("CONCAT(A1:A3)"), "#VALUE!");
  assert.equal(value('TEXTJOIN(",",TRUE,A1:A2)'), "#VALUE!");
  assert.equal(value('SUBSTITUTE("abc","b",A1)'), "text(20002)");
  assert.equal(value('SUBSTITUTE(A1,"w","ww")'), "#VALUE!");
  assert.equal(value('SUBSTITUTE("abc","","xyz")'), "text(3)", "an empty search text leaves the text unchanged, as in Excel");
});

test("uncached formula results and conditional-format display texts count against the workbook budget", () => {
  // 3,000 uncached formulas each producing 30,000 characters: 90 M code units in total.
  const cells = [ref => `<c r="${ref}" t="s"><v>0</v></c>`, ...repeat(2999, ref => `<c r="${ref}" t="str"><f>REPT(${col(0)}1,1)</f></c>`)];
  assert.throws(() => parse(workbook(cells, { shared: ["v".repeat(30_000)] })), failsWith("resource-limit", /workbook text exceeds/));
  // Conditional formats whose number format is a long literal, over 3,000 numeric cells (90 M code units).
  const styles = `<styleSheet ${S}><dxfs count="1"><dxf><numFmt numFmtId="165" formatCode="&quot;${"k".repeat(30_000)}&quot;0"/></dxf></dxfs></styleSheet>`;
  const sheet = repeat(3000, ref => `<c r="${ref}"><v>1</v></c>`);
  const bytes = workbook(sheet, { styles });
  const withRule = new Uint8Array(makeZip({
    "xl/workbook.xml": `<workbook ${S}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="st" Type="x/styles" Target="styles.xml"/></Relationships>',
    "xl/styles.xml": styles,
    "xl/worksheets/sheet1.xml": `<worksheet ${S}><sheetData>${Array.from({ length: 30 }, (_, r) => `<row r="${r + 1}">${Array.from({ length: 100 }, (_, c) => `<c r="${col(c)}${r + 1}"><v>1</v></c>`).join("")}</row>`).join("")}</sheetData><conditionalFormatting sqref="${col(0)}1:${col(99)}30"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting></worksheet>`,
  }, { deflate: true }));
  assert.ok(bytes.length > 0);
  assert.throws(() => parse(withRule), failsWith("resource-limit", /workbook text exceeds/));
});

test("legitimate workbooks with many short texts and formatted numbers stay well within the budget", () => {
  const words = Array.from({ length: 500 }, (_, index) => `Customer name ${index}`);
  const money = `<styleSheet ${S}><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00 &quot;EUR&quot;"/></numFmts><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>`;
  const cells = Array.from({ length: 100_000 }, (_, index) => index % 2 ? sharedCell(index % 500) : ref => `<c r="${ref}" s="1"><v>${index * 1000 + 0.25}</v></c>`);
  const model = parse(workbook(cells, { shared: words, styles: money }));
  assert.equal(model.sheets[0].cells.length, 100_000);
  assert.equal(model.sheets[0].cells[1].value, "Customer name 1");
  assert.match(model.sheets[0].cells[0].display, /^0\.25 EUR$/);
});
