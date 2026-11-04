import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { OdsRenderer } from "../dist/renderers/ods-renderer.js";
import { RendererRegistry, createDefaultRendererRegistry } from "../dist/renderers/registry.js";
import { detectFormat } from "../dist/detection/detect-format.js";
import { parseOds } from "../dist/parsers/odf/ods.js";
import { parseFods } from "../dist/parsers/odf/flat.js";
import { writeDemoSamples } from "../scripts/demo-samples.mjs";

// Milestone M12: ODS and OTS reach the shared workbook renderer through the default registry.
// Rendering itself needs a browser DOM and is covered by scripts/browser-regression.mjs.
test("OdsRenderer handles ods, ots and (since M15) fods", () => {
  assert.deepEqual(new OdsRenderer().formats, ["ods", "ots", "fods"]);
});

test("the default registry routes ods, ots and fods to OdsRenderer without disturbing other formats", () => {
  const registry = createDefaultRendererRegistry();
  assert.ok(registry.get("ods") instanceof OdsRenderer);
  assert.equal(registry.get("ods"), registry.get("ots"), "one renderer instance serves both formats");
  assert.equal(registry.get("fods"), registry.get("ods"), "flat ODS (M15)");
  assert.equal(registry.get("pdf"), undefined);
  for (const format of ["xlsx", "docx", "pptx", "png", "svg"]) assert.ok(registry.get(format), `${format} is still registered`);
  assert.throws(() => registry.register(new OdsRenderer()), /already registered for ods/);
  const fresh = new RendererRegistry();
  const unregister = fresh.register(new OdsRenderer());
  unregister();
  assert.equal(fresh.get("ods"), undefined);
});

test("ODS internals are not exported by the package", async () => {
  const entry = await import("../dist/index.js");
  for (const name of ["OdsRenderer", "parseOds", "renderWorkbook"]) assert.equal(name in entry, false, `${name} is internal per decision D2`);
});

test("the demo ODS sample is detected and parsed with saved display text", async () => {
  const directory = mkdtempSync(join(tmpdir(), "ufp-ods-demo-"));
  try {
    await writeDemoSamples(directory);
    const bytes = new Uint8Array(readFileSync(join(directory, "sample-spreadsheet.ods")));
    assert.equal(new TextDecoder().decode(bytes.subarray(30, 38)), "mimetype", "mimetype is the first entry");
    assert.equal(bytes[8] | (bytes[9] << 8), 0, "mimetype is stored uncompressed");
    assert.equal(detectFormat("sample-spreadsheet.ods", "", bytes), "ods");
    const workbook = parseOds(bytes, "sample-spreadsheet.ods");
    assert.deepEqual(workbook.sheets.map(sheet => sheet.name), ["Quarterly results", "Notes"]);
    assert.deepEqual(workbook.sheets[0].frozen, { rows: 1, columns: 0 }, "settings.xml freezes the header row");
    const cell = ref => workbook.sheets[0].cells.find(entry => entry.ref === ref);
    assert.deepEqual(cell("C2").data.value, { kind: "currency", lexical: "18600", number: 18600, currency: "EUR" });
    assert.equal(cell("C2").display, `18,600.00 ${String.fromCharCode(0x20ac)}`);
    assert.equal(cell("D4").display, "7.5%");
    assert.equal(cell("C6").data.value.number, 24365.5, "the saved formula result, not a recalculation");
    assert.equal(cell("C5").displayColor, "#c9211e", "negative amounts in the data style's colour");
    assert.equal(cell("C2").displayColor, undefined);
    assert.equal(workbook.styles[cell("C2").style].numberFormatInfo.supported, true);
    // The saved text agrees with the data styles, so no mismatch or raw-value notice is raised.
    assert.deepEqual(workbook.diagnostics.map(entry => entry.code), ["odf.formula.cached"]);
    // The flat demo file holds the same spreadsheet (M15).
    const flatBytes = new Uint8Array(readFileSync(join(directory, "sample-spreadsheet.fods")));
    assert.equal(detectFormat("sample-spreadsheet.fods", "", flatBytes), "fods");
    const flat = parseFods(flatBytes, "sample-spreadsheet.ods");
    assert.deepEqual(flat, { ...workbook, source: { ...workbook.source, format: "fods" } });
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});
