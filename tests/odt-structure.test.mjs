import test from "node:test";
import assert from "node:assert/strict";
import { parseOdt } from "../dist/parsers/odf/odt.js";
import { makeOdt } from "./helpers/odf.mjs";

// Milestone 3.3: lists, nested lists, tables and the styles that shape them.

const parse = (options, prefixes) => parseOdt(makeOdt(options, prefixes), "letter.odt");
const text = paragraph => paragraph.runs.map(run => run.text).join("");
const paragraphs = model => model.blocks.filter(block => block.type === "paragraph").map(block => block.value);
const tables = model => model.blocks.filter(block => block.type === "table").map(block => block.value);
const item = (body, style) => `<text:list-item><text:p${style ? ` text:style-name="${style}"` : ""}>${body}</text:p></text:list-item>`;

const numbered = '<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."><style:list-level-properties><style:list-level-label-alignment fo:margin-left="1cm" fo:text-indent="-0.5cm"/></style:list-level-properties></text:list-level-style-number><text:list-level-style-number text:level="2" style:num-format="a" style:num-suffix=")"/><text:list-level-style-number text:level="3" style:num-format="i" style:num-prefix="(" style:num-suffix=")" text:display-levels="1"/></text:list-style>';
const bulleted = '<text:list-style style:name="LB"><text:list-level-style-bullet text:level="1" text:bullet-char="-"/></text:list-style>';

test("an ordered list numbers its items and records the level", () => {
  const model = parse({ styles: numbered, body: `<text:list text:style-name="L1">${item("One")}${item("Two")}${item("Three")}</text:list>` });
  const found = paragraphs(model);
  assert.deepEqual(found.map(text), ["One", "Two", "Three"]);
  assert.deepEqual(found.map(paragraph => paragraph.list.label), ["1.", "2.", "3."]);
  assert.deepEqual(found.map(paragraph => paragraph.list.level), [0, 0, 0]);
  assert.deepEqual(found.map(paragraph => paragraph.list.kind), ["ordered", "ordered", "ordered"]);
  assert.ok(Math.abs(found[0].list.indentLeft - 28.35) < 0.1, `indentLeft was ${found[0].list.indentLeft}`);
  assert.ok(Math.abs(found[0].list.hanging - 14.17) < 0.1, `hanging was ${found[0].list.hanging}`);
});

test("nested lists number each level with its own format and restart deeper levels", () => {
  const third = `<text:list-item><text:p>Two b</text:p><text:list>${item("Deep")}</text:list></text:list-item>`;
  const body = `<text:list text:style-name="L1">${item("One")}<text:list-item><text:p>Two</text:p><text:list>${item("Two a")}${third}</text:list></text:list-item>${item("Three")}</text:list>`;
  const found = paragraphs(parse({ styles: numbered, body }));
  assert.deepEqual(found.map(text), ["One", "Two", "Two a", "Two b", "Deep", "Three"]);
  assert.deepEqual(found.map(paragraph => paragraph.list.label), ["1.", "2.", "a)", "b)", "(i)", "3."]);
  assert.deepEqual(found.map(paragraph => paragraph.list.level), [0, 0, 1, 1, 2, 0]);
});

test("a second list restarts, and one that continues the previous does not", () => {
  const first = `<text:list text:style-name="L1">${item("One")}${item("Two")}</text:list>`;
  const restarted = parse({ styles: numbered, body: `${first}<text:list text:style-name="L1">${item("Again")}</text:list>` });
  assert.deepEqual(paragraphs(restarted).map(paragraph => paragraph.list.label), ["1.", "2.", "1."]);
  const continued = parse({ styles: numbered, body: `${first}<text:list text:style-name="L1" text:continue-numbering="true">${item("Third")}</text:list>` });
  assert.deepEqual(paragraphs(continued).map(paragraph => paragraph.list.label), ["1.", "2.", "3."]);
});

test("a start value and a bullet character are honoured", () => {
  const styles = '<text:list-style style:name="LS"><text:list-level-style-number text:level="1" style:num-format="1" text:start-value="5"/></text:list-style>';
  assert.deepEqual(paragraphs(parse({ styles, body: `<text:list text:style-name="LS">${item("A")}${item("B")}</text:list>` })).map(p => p.list.label), ["5", "6"]);
  const bullets = paragraphs(parse({ styles: bulleted, body: `<text:list text:style-name="LB">${item("A")}</text:list>` }));
  assert.equal(bullets[0].list.label, "-");
  assert.equal(bullets[0].list.kind, "bullet");
});

test("a list with no style still indents and marks its items", () => {
  const found = paragraphs(parse({ body: `<text:list>${item("A")}</text:list>` }));
  assert.equal(found[0].list.kind, "bullet");
  assert.equal(found[0].list.label.length > 0, true, "a bullet is shown when the document names no style");
});

test("a list header and later paragraphs of an item are indented but not labelled", () => {
  const body = `<text:list text:style-name="L1"><text:list-header><text:p>Header</text:p></text:list-header><text:list-item><text:p>One</text:p><text:p>Still one</text:p></text:list-item></text:list>`;
  const found = paragraphs(parse({ styles: numbered, body }));
  assert.deepEqual(found.map(text), ["Header", "One", "Still one"]);
  assert.equal(found[0].list.label, "", "a header carries no number");
  assert.equal(found[1].list.label, "1.");
  assert.equal(found[2].list.label, "", "the second paragraph of an item is not numbered again");
});

test("roman and alphabetic numbering, and labels are bounded", () => {
  const styles = '<text:list-style style:name="LR"><text:list-level-style-number text:level="1" style:num-format="I"/></text:list-style>';
  const body = `<text:list text:style-name="LR">${item("a")}${item("b")}${item("c")}${item("d")}</text:list>`;
  assert.deepEqual(paragraphs(parse({ styles, body })).map(p => p.list.label), ["I", "II", "III", "IV"]);
  const huge = '<text:list-style style:name="LH"><text:list-level-style-number text:level="1" style:num-format="1" style:num-prefix="' + "x".repeat(200) + '"/></text:list-style>';
  const label = paragraphs(parse({ styles: huge, body: `<text:list text:style-name="LH">${item("a")}</text:list>` }))[0].list.label;
  assert.ok(label.length <= 64, `label was ${label.length} characters`);
});

test("a table becomes rows of cells with their paragraphs", () => {
  const body = '<table:table table:name="T"><table:table-row><table:table-cell><text:p>A1</text:p></table:table-cell><table:table-cell><text:p>B1</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell><text:p>A2</text:p></table:table-cell><table:table-cell><text:p>B2</text:p></table:table-cell></table:table-row></table:table>';
  const [table] = tables(parse({ body }));
  assert.equal(table.rows.length, 2);
  assert.deepEqual(table.rows.map(row => row.cells.map(cell => text(cell.paragraphs[0]))), [["A1", "B1"], ["A2", "B2"]]);
  assert.deepEqual(table.rows[0].cells.map(cell => cell.column), [0, 1]);
});

test("spanned and covered cells keep the grid aligned", () => {
  const body = '<table:table table:name="T"><table:table-row><table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="2"><text:p>wide</text:p></table:table-cell><table:covered-table-cell/><table:table-cell><text:p>C1</text:p></table:table-cell></table:table-row></table:table>';
  const [table] = tables(parse({ body }));
  const [first, second] = table.rows[0].cells;
  assert.equal(first.colSpan, 2);
  assert.equal(first.rowSpan, 2);
  assert.equal(second.column, 3, "the covered cell and the span both advance the column");
  assert.equal(text(second.paragraphs[0]), "C1");
});

test("repeated columns, rows and cells are expanded within bounds", () => {
  const body = '<table:table table:name="T"><table:table-column table:number-columns-repeated="3" table:style-name="co1"/><table:table-row table:number-rows-repeated="2"><table:table-cell table:number-columns-repeated="3"><text:p>x</text:p></table:table-cell></table:table-row></table:table>';
  const automatic = '<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="2cm"/></style:style>';
  const [table] = tables(parse({ automatic, body }));
  assert.equal(table.rows.length, 2);
  assert.equal(table.rows[0].cells.length, 3);
  assert.ok(Math.abs(table.rows[0].cells[0].width - 56.69) < 0.1, `width was ${table.rows[0].cells[0].width}`);
  // A repeat count decides how many objects are built, so it must be clamped rather than trusted.
  const bomb = '<table:table table:name="T"><table:table-row table:number-rows-repeated="99999999"><table:table-cell><text:p>x</text:p></table:table-cell></table:table-row></table:table>';
  assert.ok(tables(parse({ body: bomb }))[0].rows.length <= 10_000);
});

test("table and cell styles are applied", () => {
  const automatic = [
    '<style:style style:name="Ta1" style:family="table"><style:table-properties style:width="12cm" style:rel-width="80%" table:align="center"/></style:style>',
    '<style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:background-color="#ffeeee" style:vertical-align="middle" fo:border="0.5pt solid #333333" fo:padding="2pt"/></style:style>',
  ].join("");
  const body = '<table:table table:name="T" table:style-name="Ta1"><table:table-row><table:table-cell table:style-name="ce1"><text:p>A</text:p></table:table-cell></table:table-row></table:table>';
  const [table] = tables(parse({ automatic, body }));
  assert.ok(Math.abs(table.width - 340.16) < 0.5, `width was ${table.width}`);
  assert.equal(table.widthPercent, 80);
  assert.equal(table.alignment, "center");
  const cell = table.rows[0].cells[0];
  assert.equal(cell.background, "#ffeeee");
  assert.equal(cell.verticalAlign, "center");
  assert.equal(cell.borders.top.color, "#333333");
  assert.ok(Math.abs(cell.margins.top - 2) < 0.01);
});

test("a list inside a cell keeps its labels, and a nested table is flattened with a report", () => {
  const body = `<table:table table:name="T"><table:table-row><table:table-cell><text:list text:style-name="L1">${item("One")}${item("Two")}</text:list></table:table-cell><table:table-cell><table:table table:name="Inner"><table:table-row><table:table-cell><text:p>inner</text:p></table:table-cell></table:table-row></table:table></table:table-cell></table:table-row></table:table>`;
  const model = parse({ styles: numbered, body });
  const [table] = tables(model);
  const [listCell, nestedCell] = table.rows[0].cells;
  assert.deepEqual(listCell.paragraphs.map(text), ["One", "Two"]);
  assert.deepEqual(listCell.paragraphs.map(paragraph => paragraph.list.label), ["1.", "2."]);
  assert.deepEqual(nestedCell.paragraphs.map(text), ["inner"], "the inner table's text is kept");
  assert.ok(model.diagnostics?.some(entry => entry.severity === "warning"), "losing the inner table's shape is reported");
});

test("blocks keep document order across paragraphs, lists and tables", () => {
  const body = `<text:p>Intro</text:p><text:list text:style-name="L1">${item("Item")}</text:list><table:table table:name="T"><table:table-row><table:table-cell><text:p>Cell</text:p></table:table-cell></table:table-row></table:table><text:p>Outro</text:p>`;
  const model = parse({ styles: numbered, body });
  assert.deepEqual(model.blocks.map(block => block.type), ["paragraph", "paragraph", "table", "paragraph"]);
  assert.deepEqual(paragraphs(model).map(text), ["Intro", "Item", "Outro"]);
});

test("an empty cell still holds a paragraph so the grid keeps its shape", () => {
  const body = '<table:table table:name="T"><table:table-row><table:table-cell/><table:table-cell><text:p>B</text:p></table:table-cell></table:table-row></table:table>';
  const [table] = tables(parse({ body }));
  assert.equal(table.rows[0].cells.length, 2);
  assert.deepEqual(table.rows[0].cells[0].paragraphs.length, 1);
  assert.equal(text(table.rows[0].cells[0].paragraphs[0]), "");
});

test("lists and tables work with renamed prefixes", () => {
  const renamed = parse({ styles: numbered, body: `<text:list text:style-name="L1">${item("One")}</text:list><table:table table:name="T"><table:table-row><table:table-cell><text:p>A</text:p></table:table-cell></table:table-row></table:table>` }, { office: "o", text: "x", table: "tb", style: "s", fo: "f", meta: "m" });
  assert.equal(paragraphs(renamed)[0].list.label, "1.");
  assert.equal(text(tables(renamed)[0].rows[0].cells[0].paragraphs[0]), "A");
});
