import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Worker } from "node:worker_threads";
import { PreviewError } from "../dist/errors.js";
import { parseModel } from "../dist/parsers/parse-model.js";
import { installParserWorker } from "../dist/worker/worker-main.js";
import { prepareTransfer } from "../dist/worker/transfer.js";
import { WorkerClient } from "../dist/worker/worker-client.js";
import { isBuiltInDocumentRenderer, createDefaultRendererRegistry } from "../dist/renderers/registry.js";
import { writeDemoSamples } from "../scripts/demo-samples.mjs";
import { buildFixture, ooxmlFixtures } from "./helpers/ooxml-fixtures.mjs";
import { makeOds, makeFods, makeOdt, OTT_MEDIA_TYPE } from "./helpers/odf.mjs";
import { securityFixtures } from "./helpers/security-fixtures.mjs";

// The parser worker (Phase 2 step 6): the worker logic with a fake scope, the transfer preparation, and the
// real thing in a Node worker thread driven through the viewer's WorkerClient.

const workerUrl = new URL("./helpers/node-parser-worker.mjs", import.meta.url);

/** The Web Worker interface the client uses, over a Node worker thread. */
class NodeWorker {
  #worker; #listeners = new Map();
  constructor(options) { this.#worker = new Worker(workerUrl, options); }
  addEventListener(type, listener) { const wrapped = type === "message" ? data => listener({ data }) : error => listener({ type, error }); this.#listeners.set(listener, [type, wrapped]); this.#worker.on(type, wrapped); }
  removeEventListener(type, listener) { const entry = this.#listeners.get(listener); if (entry) { this.#worker.off(entry[0], entry[1]); this.#listeners.delete(listener); } }
  postMessage(data, transfer) { this.#worker.postMessage(data, transfer); }
  terminate() { void this.#worker.terminate(); }
}

/** Every model the tests can build: OOXML fixtures, demo samples, an ODS and a FODS workbook, and the security fixtures with images. */
async function documents() {
  const list = ooxmlFixtures.filter(fixture => !fixture.options).map(fixture => ({ id: fixture.id, format: fixture.kind, bytes: new Uint8Array(buildFixture(fixture)), name: fixture.file }));
  const directory = mkdtempSync(join(tmpdir(), "ufp-worker-"));
  try {
    await writeDemoSamples(directory);
    for (const [name, format] of [["sample-report.docx", "docx"], ["sample-workbook.xlsx", "xlsx"], ["sample-deck.pptx", "pptx"]]) list.push({ id: name, format, bytes: new Uint8Array(readFileSync(join(directory, name))), name });
  } finally { rmSync(directory, { recursive: true, force: true }); }
  const body = '<table:table table:name="S"><table:table-row><table:table-cell office:value-type="string"><text:p>ods</text:p></table:table-cell></table:table-row></table:table>';
  list.push({ id: "ods", format: "ods", bytes: makeOds({ body }), name: "s.ods" }, { id: "fods", format: "fods", bytes: makeFods({ body }), name: "s.fods" });
  const textBody = '<text:h text:outline-level="1">Title</text:h><text:p>Body <text:span text:style-name="T1">bold</text:span></text:p>';
  const automatic = '<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>';
  list.push({ id: "odt", format: "odt", bytes: makeOdt({ body: textBody, automatic }), name: "letter.odt" }, { id: "ott", format: "ott", bytes: makeOdt({ body: textBody, automatic, mimetype: OTT_MEDIA_TYPE }), name: "letter.ott" });
  for (const [name, fixture] of Object.entries(securityFixtures("http://attacker.invalid:9"))) if (["pptx", "xlsx", "docx"].includes(fixture.kind)) list.push({ id: name, format: fixture.kind, bytes: new Uint8Array(fixture.bytes), name });
  return list;
}
const mainThread = entry => parseModel(entry.format, entry.bytes, entry.name, { locale: "en-US" });
/** All Uint8Arrays in a model. */
function views(model) {
  const found = [], stack = [model];
  while (stack.length) { const value = stack.pop(); if (value instanceof Uint8Array) found.push(value); else if (value && typeof value === "object") stack.push(...Object.values(value)); }
  return found;
}

test("prepareTransfer copies views that do not cover exactly their own buffer, and never transfers the input", () => {
  const input = new ArrayBuffer(16), inside = new Uint8Array(input, 4, 4).fill(7), whole = new Uint8Array(input);
  const own = new Uint8Array([1, 2, 3]), shorter = new Uint8Array(new ArrayBuffer(10), 0, 3);
  const model = { a: { image: { bytes: inside } }, b: [{ bytes: inside }], c: { bytes: whole }, d: { bytes: own }, e: { bytes: shorter }, cells: [{ ref: "A1", value: "1" }] };
  const transfer = prepareTransfer(model, input);
  assert.notEqual(model.a.image.bytes, inside);
  assert.equal(model.a.image.bytes, model.b[0].bytes, "a view placed twice is copied once and stays shared");
  assert.deepEqual([...model.a.image.bytes], [7, 7, 7, 7]);
  assert.notEqual(model.c.bytes.buffer, input, "a view of the whole input is copied too");
  assert.equal(model.d.bytes, own, "a view of its own whole buffer is kept");
  assert.equal(model.e.bytes.byteLength, 3); assert.equal(model.e.bytes.buffer.byteLength, 3);
  assert.equal(transfer.length, 4);
  assert.ok(!transfer.includes(input));
  for (const view of views(model)) assert.ok(view.byteOffset === 0 && view.buffer.byteLength === view.byteLength && transfer.includes(view.buffer));
  assert.deepEqual(inside.byteLength, 4, "the original view is left as it was");
});

test("the worker reports ready first, then answers each request with a result or a bounded error", () => {
  const posted = [];
  let listener;
  installParserWorker({ addEventListener(type, handler) { listener = handler; }, postMessage(message, transfer) { posted.push({ message, transfer }); } });
  assert.deepEqual(posted[0].message, { v: 1, type: "ready" });
  const fixture = ooxmlFixtures.find(entry => entry.id === "docx-basic"), bytes = new Uint8Array(buildFixture(fixture));
  listener({ data: { v: 1, type: "parse", id: 1, format: "docx", name: "a.docx", bytes: bytes.slice().buffer } });
  assert.equal(posted[1].message.type, "result"); assert.equal(posted[1].message.id, 1);
  assert.deepEqual(structuredClone(posted[1].message.model), structuredClone(mainThread({ format: "docx", bytes, name: "a.docx" })));
  // A file the parser refuses: the same code and message as in main-thread mode.
  const corrupt = new TextEncoder().encode("not a zip");
  let mainError; try { mainThread({ format: "xlsx", bytes: corrupt, name: "x.xlsx" }); } catch (error) { mainError = error; }
  listener({ data: { v: 1, type: "parse", id: 2, format: "xlsx", name: "x.xlsx", bytes: corrupt.slice().buffer } });
  assert.deepEqual(posted[2].message, { v: 1, type: "error", id: 2, error: { code: mainError.code, message: mainError.message } });
  // An invalid request with a usable id is answered with renderer-error; without one it is ignored.
  listener({ data: { v: 1, type: "parse", id: 3, format: "pdf", name: "x", bytes: new ArrayBuffer(1) } });
  assert.deepEqual(posted[3].message, { v: 1, type: "error", id: 3, error: { code: "renderer-error", message: "The parser worker received an invalid request." } });
  for (const data of [null, "parse", { v: 1, type: "parse" }, { id: -1 }]) listener({ data });
  assert.equal(posted.length, 4);
});

test("in a worker thread, every model reaches the client intact, validated, with each image in a buffer of its own", async t => {
  const client = new WorkerClient({ create: () => new NodeWorker(), maxFileSize: 64 * 1024 * 1024 });
  t.after(() => client.destroy());
  let images = 0;
  for (const entry of await documents()) {
    const model = await client.parse(entry.format, entry.bytes, entry.name, { locale: "en-US" });
    assert.deepStrictEqual(model, structuredClone(mainThread(entry)), entry.id);
    for (const view of views(model)) { images++; assert.ok(view.byteOffset === 0 && view.buffer.byteLength === view.byteLength, `${entry.id}: a received view shares a larger buffer`); }
  }
  assert.ok(images >= 5, `only ${images} images crossed the thread boundary`);
  // Errors cross the boundary with the same code and message as in main-thread mode.
  const corrupt = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3]);
  let mainError; try { mainThread({ format: "pptx", bytes: corrupt, name: "x.pptx" }); } catch (error) { mainError = error; }
  await assert.rejects(client.parse("pptx", corrupt, "x.pptx"), error => error instanceof PreviewError && error.code === mainError.code && error.message === mainError.message);
  // The caller's input was copied, not transferred.
  assert.ok(corrupt.byteLength === 7);
});

test("in a worker thread, transferred buffers are detached on the worker side after sending", async t => {
  const worker = new Worker(workerUrl, { workerData: { report: true } });
  t.after(() => worker.terminate());
  const messages = [], waiting = [];
  worker.on("message", message => { if (waiting.length) waiting.shift()(message); else messages.push(message); });
  const next = () => new Promise(resolve => { if (messages.length) resolve(messages.shift()); else waiting.push(resolve); });
  const ready = await next();
  assert.deepEqual(ready, { v: 1, type: "ready" });
  // The two OOXML fixtures with embedded images.
  for (const id of ["xlsx-table", "xlsx-slicer"]) {
    const fixture = ooxmlFixtures.find(entry => entry.id === id), bytes = new Uint8Array(buildFixture(fixture));
    const request = { v: 1, type: "parse", id: id.length, format: "xlsx", name: fixture.file, bytes: bytes.slice().buffer };
    worker.postMessage(request, [request.bytes]);
    assert.equal(request.bytes.byteLength, 0, "the request buffer is transferred to the worker");
    const result = await next(), report = await next();
    assert.equal(result.type, "result");
    const received = views(result.model);
    assert.ok(received.length >= 1, `${id}: no image`);
    assert.equal(report.transferred, new Set(received.map(view => view.buffer)).size, `${id}: one transfer per distinct buffer`);
    assert.ok(report.lengthsAfterSending.every(length => length === 0), `${id}: buffers still attached in the worker: ${report.lengthsAfterSending}`);
    const local = views(mainThread({ format: "xlsx", bytes, name: fixture.file }));
    assert.deepEqual(received.map(view => [...view]), local.map(view => [...view]), `${id}: image bytes differ`);
  }
});

test("only the library's document renderers render worker models; renderParsed refuses a model of another kind", () => {
  const registry = createDefaultRendererRegistry();
  for (const format of ["docx", "xlsx", "pptx", "ods", "ots", "fods", "odt", "ott"]) assert.equal(isBuiltInDocumentRenderer(registry.get(format)), true, format);
  assert.equal(isBuiltInDocumentRenderer(registry.get("png")), false);
  const custom = { formats: ["docx"], render() { return { element: {} }; }, renderParsed() { return { element: {} }; } };
  assert.equal(isBuiltInDocumentRenderer(custom), false, "a custom renderer is never given a worker model");
  const word = { kind: "word", title: "t", blocks: [] }, workbook = { kind: "workbook", title: "t", sheets: [] };
  const context = { file: { name: "x" }, format: "docx", bytes: new Uint8Array(), signal: new AbortController().signal };
  assert.throws(() => registry.get("docx").renderParsed(workbook, context), error => error.code === "renderer-error");
  assert.throws(() => registry.get("pptx").renderParsed(word, context), error => error.code === "renderer-error");
  assert.throws(() => registry.get("xlsx").renderParsed(word, context), error => error.code === "renderer-error");
  assert.throws(() => registry.get("ods").renderParsed(word, context), error => error.code === "renderer-error");
});

test("the worker entry point is exported and declared as having side effects, so bundlers keep it", () => {
  const manifest = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8"));
  assert.deepEqual(manifest.exports["./worker"], { types: "./dist/worker.d.ts", import: "./dist/worker.js" });
  // Without this, esbuild, Rollup and webpack drop `import "@dukeofwest/universal-file-preview/worker"` and emit an
  // empty worker (seen with the Angular application builder before this entry was added).
  assert.ok(manifest.sideEffects.includes("./dist/worker.js"), JSON.stringify(manifest.sideEffects));
  assert.ok(manifest.sideEffects.includes("**/*.css"));
});
