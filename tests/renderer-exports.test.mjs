import test from "node:test";
import assert from "node:assert/strict";

// Milestone M9a: workbook rendering is separate from XLSX parsing. Rendering itself needs a
// browser DOM and is covered by scripts/browser-regression.mjs; this checks the module seam.
test("renderWorkbook is exported by the spreadsheet renderer module but not by the package", async () => {
  const renderer = await import("../dist/renderers/xlsx-renderer.js");
  assert.equal(typeof renderer.renderWorkbook, "function");
  assert.equal(renderer.renderWorkbook.length, 2, "renderWorkbook(workbook, options)");
  assert.deepEqual(new renderer.XlsxRenderer().formats, ["xlsx"]);
  const entry = await import("../dist/index.js");
  assert.equal("renderWorkbook" in entry, false, "internal per decision D2");
});
