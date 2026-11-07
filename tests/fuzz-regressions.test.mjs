import test from "node:test";
import assert from "node:assert/strict";
import { parseDocx, parseXlsx } from "../dist/parsers/ooxml.js";
import { columnName } from "../dist/parsers/xlsx.js";
import { ooxmlFixtures } from "./helpers/ooxml-fixtures.mjs";
import { makeZip } from "./helpers/zip.mjs";

// Inputs found by scripts/fuzz.mjs (seed 1), reduced to the attribute that triggered each bug.

/** A fixture's package with one part's text changed. */
function variant(id, path, change) {
  const entries = ooxmlFixtures.find(fixture => fixture.id === id).entries(), text = String(entries[path]);
  const changed = change(text);
  assert.notEqual(changed, text, `${id}: the change did not apply`);
  return new Uint8Array(makeZip({ ...entries, [path]: changed }));
}

test("a list level that is not a whole number no longer breaks DOCX numbering (RangeError: Invalid array length)", () => {
  for (const value of ["NaN", "1.5", "-3", "Infinity", "1e309", "A1:XFD1048576", ""]) {
    const bytes = variant("docx-numbered", "word/document.xml", text => text.replace(/(<w:ilvl w:val=")[^"]*(")/, `$1${value}$2`));
    const model = parseDocx(bytes, "numbered.docx");
    const levels = model.blocks.flatMap(block => block.type === "paragraph" && block.value.list ? [block.value.list.level] : []);
    assert.ok(levels.every(level => Number.isInteger(level) && level >= 0 && level <= 8), `${value}: levels ${levels}`);
  }
});

test("a filter column index outside its table no longer sends the slicer into an endless loop (out of memory)", () => {
  for (const value of ["Infinity", "-Infinity", "1e9", "-1", "0.5", "NaN"]) {
    const bytes = variant("xlsx-slicer", "xl/tables/table1.xml", text => text.replace(/colId="[^"]*"/, `colId="${value}"`));
    const started = performance.now(), model = parseXlsx(bytes, "slicer.xlsx");
    assert.ok(performance.now() - started < 2000, value);
    assert.ok(model.sheets.flatMap(sheet => sheet.drawings ?? []).some(drawing => drawing.kind === "slicer"), `${value}: the slicer is still shown`);
  }
  for (const value of [Infinity, -Infinity, NaN, 1.5, 2 ** 60]) assert.equal(columnName(value), "", String(value));
  assert.equal(columnName(1), "A"); assert.equal(columnName(28), "AB"); assert.equal(columnName(16_384), "XFD");
});
