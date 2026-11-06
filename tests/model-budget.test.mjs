import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Worker } from "node:worker_threads";
import { checkModel, MODEL_LIMITS, ModelBudget, modelBackstopRejections, modelBudgetMisses } from "../dist/model-budget.js";
import { parseDocx, parsePptx, parseXlsx } from "../dist/parsers/ooxml.js";
import { parseOds } from "../dist/parsers/odf/ods.js";
import { writeDemoSamples } from "../scripts/demo-samples.mjs";
import { buildFixture, ooxmlFixtures } from "./helpers/ooxml-fixtures.mjs";
import { makeOds } from "./helpers/odf.mjs";
import { makeZip } from "./helpers/zip.mjs";

// Output accounting (model-budget.ts): parsers charge a ModelBudget while they build the model, and checkModel
// walks the finished model as a backstop. Strings count per occurrence; typed arrays and objects that contain
// objects count once per identity (as structured cloning copies them); leaf objects count per reference.

const failsWith = (code, pattern) => error => error?.name === "PreviewError" && error.code === code && pattern.test(error.message);
const parsers = { docx: parseDocx, pptx: parsePptx, xlsx: parseXlsx, ods: parseOds };
const small = { ...MODEL_LIMITS, objects: 10, stringUnits: 100, stringLength: 50, binaryBytes: 16, depth: 4 };

/** Runs `work` and asserts that the backstop neither rejected a model nor measured more than a parser charged. */
function withoutBackstop(work) {
  const rejections = modelBackstopRejections(), misses = modelBudgetMisses();
  try { return work(); }
  finally {
    assert.equal(modelBackstopRejections(), rejections, "the backstop walk rejected a model the parser's budget accepted");
    assert.equal(modelBudgetMisses(), misses, "the finished model measured more than the parser charged");
  }
}

test("strings count once per occurrence, also when the same string is referenced several times", () => {
  const text = "x".repeat(10);
  assert.equal(checkModel({ a: text, b: text, c: [text, text] }).stringUnits, 40);
  // The case measured for the design: cloning writes each occurrence, so 100,000 references to one
  // 1,000,000-character string are 10^11 code units and are refused, although the heap holds the string once.
  const references = Array.from({ length: 100_000 }, () => "y".repeat(1_000_000).slice(0));
  assert.throws(() => checkModel(references), failsWith("resource-limit", /text limit/));
});

test("objects that contain objects count once per identity; leaf objects count per reference", () => {
  const shared = { inner: { value: 1 } };
  // array + shared (once) + inner (once, reached through shared only once)
  assert.equal(checkModel([shared, shared, shared]).objects, 3);
  const leaf = { value: 1 };
  // array + three references to the leaf (counted per reference: an overestimate, never an underestimate)
  assert.equal(checkModel([leaf, leaf, leaf]).objects, 4);
  // A diamond is a shared acyclic subtree and is allowed.
  const bottom = { list: [1, 2] }, left = { bottom }, right = { bottom };
  assert.equal(checkModel({ left, right }).objects, 5);
});

test("a Uint8Array counts once per identity, as cloning copies it once", () => {
  const bytes = new Uint8Array(1000);
  assert.equal(checkModel([{ bytes }, { bytes }]).binaryBytes, 1000);
  assert.equal(checkModel([{ bytes }, { bytes: new Uint8Array(1000) }]).binaryBytes, 2000);
});

test("cycles are rejected, directly or through arrays and shared subtrees", () => {
  const self = { name: "a" }; self.self = self;
  assert.throws(() => checkModel(self), failsWith("invalid-file", /cycle/));
  const parent = { children: [] }, childNode = { parent, list: [1] }; parent.children.push(childNode);
  assert.throws(() => checkModel({ root: parent }), failsWith("invalid-file", /cycle/));
  const array = []; array.push({ back: array });
  assert.throws(() => checkModel(array), failsWith("invalid-file", /cycle/));
});

test("only plain data is accepted: no functions, symbols, big integers, non-finite numbers or class instances", () => {
  for (const value of [() => 1, Symbol("s"), 1n, Number.NaN, Infinity, new Date(0), new Map(), new Set(), new (class Custom {})(), new Uint16Array(2)]) {
    assert.throws(() => checkModel({ value }), failsWith("invalid-file", /./), String(typeof value));
  }
  assert.doesNotThrow(() => checkModel({ a: null, b: undefined, c: true, d: 0, e: "", f: [], g: Object.create(null) }));
});

test("each limit is enforced by the walk and by the budget, and a walk rejection is counted", () => {
  const before = modelBackstopRejections();
  assert.throws(() => checkModel(Array.from({ length: 11 }, () => ({})), small), failsWith("resource-limit", /object limit/));
  assert.throws(() => checkModel(["a".repeat(51)], small), failsWith("resource-limit", /string length limit/));
  assert.throws(() => checkModel(["a".repeat(40), "b".repeat(40), "c".repeat(40)], small), failsWith("resource-limit", /text limit/));
  assert.throws(() => checkModel([new Uint8Array(17)], small), failsWith("resource-limit", /binary data limit/));
  assert.throws(() => checkModel({ a: { b: { c: { d: { e: 1 } } } } }, small), failsWith("resource-limit", /nesting depth/));
  assert.equal(modelBackstopRejections(), before + 5);

  const budget = new ModelBudget(small);
  assert.throws(() => { for (let index = 0; index < 11; index++) budget.objects(); }, failsWith("resource-limit", /object limit/));
  assert.throws(() => new ModelBudget(small).text("a".repeat(51)), failsWith("resource-limit", /string length/));
  assert.throws(() => { const b = new ModelBudget(small); b.text("a".repeat(50)); b.text("a".repeat(50)); b.text("a"); }, failsWith("resource-limit", /text limit/));
  assert.throws(() => new ModelBudget(small).binary(new Uint8Array(17)), failsWith("resource-limit", /binary data/));
  assert.throws(() => new ModelBudget({ ...small, cells: 2 }).cells(3), failsWith("resource-limit", /more than 2 cells/));
  // The same Uint8Array placed twice is charged once.
  const bytes = new Uint8Array(10), twice = new ModelBudget(small);
  twice.binary(bytes); twice.binary(bytes);
  assert.equal(twice.usage.binaryBytes, 10);
});

test("the default limits are 500,000 cells, 6,000,000 objects, 64 Mi string code units and 256 MiB of binary data", () => {
  assert.equal(MODEL_LIMITS.cells, 500_000);
  assert.equal(MODEL_LIMITS.objects, 6_000_000);
  assert.equal(MODEL_LIMITS.stringUnits, 64 * 1024 * 1024);
  assert.equal(MODEL_LIMITS.binaryBytes, 256 * 1024 * 1024);
});

/** Every model the tests can build from fixtures: the OOXML fixtures, the demo samples and an ODS workbook. */
async function fixtureModels() {
  const models = ooxmlFixtures.map(fixture => ({ id: fixture.id, kind: fixture.kind, bytes: buildFixture(fixture), name: fixture.file, options: fixture.options }));
  const directory = mkdtempSync(join(tmpdir(), "ufp-budget-"));
  try {
    await writeDemoSamples(directory);
    for (const [file, kind] of [["sample-report.docx", "docx"], ["sample-workbook.xlsx", "xlsx"], ["sample-deck.pptx", "pptx"]]) models.push({ id: file, kind, bytes: new Uint8Array(readFileSync(join(directory, file))), name: file });
  } finally { rmSync(directory, { recursive: true, force: true }); }
  models.push({ id: "ods", kind: "ods", name: "s.ods", bytes: makeOds({ body: '<table:table table:name="S"><table:table-row><table:table-cell office:value-type="string" table:number-columns-repeated="3"><text:p>shared <text:span>text</text:span></text:p></table:table-cell><table:table-cell office:value-type="float" office:value="1.5"><text:p>1.5</text:p><office:annotation><text:p>note</text:p></office:annotation></table:table-cell></table:table-row></table:table>' }) });
  return models;
}

test("every fixture model is charged completely while it is parsed: the backstop measures no more than the parser charged", async () => {
  for (const fixture of await fixtureModels()) {
    const model = withoutBackstop(() => parsers[fixture.kind](fixture.bytes, fixture.name, fixture.options));
    // ModelBudget.add applies the walk's rules, so charging a whole finished model gives exactly what the walk measures.
    const budget = new ModelBudget();
    budget.add(model);
    const usage = checkModel(model), charged = budget.usage;
    assert.deepEqual({ objects: charged.objects, stringUnits: charged.stringUnits, binaryBytes: charged.binaryBytes }, { objects: usage.objects, stringUnits: usage.stringUnits, binaryBytes: usage.binaryBytes }, fixture.id);
  }
});

const S = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
/** An XLSX workbook whose cells all use one long shared string (each reference is within the per-cell limit). */
function sharedStringWorkbook(references) {
  const column = index => String.fromCharCode(65 + Math.floor(index / 26)) + String.fromCharCode(65 + index % 26), rows = [];
  for (let start = 0; start < references; start += 100) rows.push(`<row r="${start / 100 + 1}">${Array.from({ length: Math.min(100, references - start) }, (_, index) => `<c r="${column(index)}${start / 100 + 1}" t="s"><v>0</v></c>`).join("")}</row>`);
  return new Uint8Array(makeZip({
    "xl/workbook.xml": `<workbook ${S}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="ss" Type="x/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
    "xl/sharedStrings.xml": `<sst ${S}><si><t>${"s".repeat(32_767)}</t></si></sst>`,
    "xl/worksheets/sheet1.xml": `<worksheet ${S}><sheetData>${rows.join("")}</sheetData></worksheet>`,
  }, { deflate: true }));
}
/** An ODS sheet of one row of repeated cells sharing one long text. */
const repeatedTextOds = (copies, rows = 1) => makeOds({ body: `<table:table table:name="S"><table:table-row table:number-rows-repeated="${rows}"><table:table-cell office:value-type="string" table:number-columns-repeated="${copies}"><text:p>${"t".repeat(32_767)}</text:p></table:table-cell></table:table-row></table:table>` });

test("crafted oversized models are stopped by the parsers' budgets, not by the backstop", () => {
  withoutBackstop(() => {
    // XLSX: one shared string referenced by 5,000 cells (value and display text per cell).
    assert.throws(() => parseXlsx(sharedStringWorkbook(5000), "x.xlsx"), failsWith("resource-limit", /text limit/));
    // ODS: repeated cells share their data object but each copy carries its own value and display strings.
    assert.throws(() => parseOds(repeatedTextOds(3000), "x.ods"), failsWith("resource-limit", /text limit/));
    // ODS: the cell limit (500,000 by default) is checked before the copies are materialised.
    const cells = (columns, rows) => makeOds({ body: `<table:table table:name="S"><table:table-row table:number-rows-repeated="${rows}"><table:table-cell office:value-type="float" office:value="1" table:number-columns-repeated="${columns}"/></table:table-row></table:table>` });
    assert.throws(() => parseOds(cells(501, 1000), "x.ods"), failsWith("resource-limit", /cell limit/));
  });
});

test("an ODS workbook with exactly 500,000 cells is accepted by default; callers can raise maxCells", () => {
  const cells = makeOds({ body: '<table:table table:name="S"><table:table-row table:number-rows-repeated="1000"><table:table-cell office:value-type="float" office:value="1" table:number-columns-repeated="500"/></table:table-row></table:table>' });
  const model = withoutBackstop(() => parseOds(cells, "x.ods"));
  assert.equal(model.sheets[0].cells.length, 500_000);
  const larger = makeOds({ body: '<table:table table:name="S"><table:table-row table:number-rows-repeated="1000"><table:table-cell office:value-type="float" office:value="1" table:number-columns-repeated="501"/></table:table-row></table:table>' });
  assert.equal(withoutBackstop(() => parseOds(larger, "x.ods", { limits: { maxCells: 501_000 } })).sheets[0].cells.length, 501_000);
});

/** One worker thread for the tests below; requests are matched by id. */
function startWorker() {
  const worker = new Worker(new URL("./helpers/model-worker.mjs", import.meta.url));
  let next = 0;
  const pending = new Map();
  worker.on("message", message => { pending.get(message.id)?.(message); pending.delete(message.id); });
  return {
    parse(kind, bytes, name) { return new Promise(resolve => { const id = next++; pending.set(id, resolve); worker.postMessage({ id, kind, bytes, name }); }); },
    close: () => worker.terminate(),
  };
}

test("in a worker thread, the same models are produced, cloned to the main thread intact, and measure the same there", async () => {
  const worker = startWorker();
  try {
    for (const fixture of await fixtureModels()) {
      if (fixture.options) continue; // the worker parses with default options
      const local = parsers[fixture.kind](fixture.bytes, fixture.name);
      const reply = await worker.parse(fixture.kind, fixture.bytes, fixture.name);
      assert.equal(reply.error, undefined, `${fixture.id}: ${reply.error?.message}`);
      assert.equal(reply.rejections, 0, `${fixture.id}: backstop rejections in the worker`);
      assert.equal(reply.misses, 0, `${fixture.id}: budget misses in the worker`);
      // The worker's model arrives as a structured clone of the same model (Node Buffers, which the test zip helper
      // produces for stored images, arrive as plain Uint8Arrays with the same bytes, as in browsers).
      assert.deepStrictEqual(reply.model, structuredClone(local), fixture.id);
      // The walk on the main thread measures the received model exactly as the worker measured the original.
      assert.deepEqual(withoutBackstop(() => checkModel(reply.model)), reply.usage, fixture.id);
    }
  } finally { await worker.close(); }
});

test("in a worker thread, crafted oversized models are stopped by the budget with the same errors", async () => {
  const worker = startWorker();
  try {
    for (const [kind, bytes, pattern] of [["xlsx", sharedStringWorkbook(5000), /text limit/], ["ods", repeatedTextOds(3000), /text limit/]]) {
      const reply = await worker.parse(kind, bytes, `x.${kind}`);
      assert.equal(reply.model, undefined);
      assert.equal(reply.error.name, "PreviewError");
      assert.equal(reply.error.code, "resource-limit");
      assert.match(reply.error.message, pattern);
      assert.equal(reply.rejections, 0, "the budget, not the backstop, stopped the model");
    }
  } finally { await worker.close(); }
});

test("structured cloning keeps shared subtrees shared and copies each string occurrence, as the accounting assumes", () => {
  const sharedRuns = [{ text: "run", font: { bold: true } }];
  const model = { cells: [{ ref: "A1", runs: sharedRuns }, { ref: "A2", runs: sharedRuns }] };
  const copy = structuredClone(model);
  assert.equal(copy.cells[0].runs, copy.cells[1].runs, "one shared array arrives as one array");
  assert.equal(checkModel(copy).objects, checkModel(model).objects);
  // Cycles survive structuredClone, but are not valid models: the walk rejects them before cloning.
  const cyclic = { a: {} }; cyclic.a.back = cyclic;
  assert.doesNotThrow(() => structuredClone(cyclic));
  assert.throws(() => checkModel(cyclic), failsWith("invalid-file", /cycle/));
});
