import test from "node:test";
import assert from "node:assert/strict";
import { FilePreviewer, PreviewError, RendererRegistry } from "../dist/index.js";
import { DocxRenderer } from "../dist/renderers/ooxml-renderers.js";
import { WorkerStartupError } from "../dist/worker/worker-client.js";
import { buildFixture, ooxmlFixtures } from "./helpers/ooxml-fixtures.mjs";

// FilePreviewer in worker mode (Phase 2 step 7) with a minimal fake DOM, a fake worker and test renderers:
// which path renders, fallback, timeouts, cancellation by signal, new load and destroy(), and the
// load-generation guard. Real browsers are covered by scripts/browser-security.mjs and the regression script.

class FakeElement {
  constructor(tag) { this.tagName = tag; this.children = []; this.attributes = {}; this.className = ""; this.textContent = ""; this.title = ""; this.hidden = false; this.parent = null; this.style = { setProperty() {} }; }
  setAttribute(name, value) { this.attributes[name] = String(value); }
  append(...nodes) { for (const node of nodes) { node.remove?.(); node.parent = this; this.children.push(node); } }
  replaceChildren(...nodes) { for (const child of this.children) child.parent = null; this.children = []; this.append(...nodes); }
  remove() { if (this.parent) this.parent.children = this.parent.children.filter(child => child !== this); this.parent = null; }
  addEventListener() {}
  dispatchEvent() {}
  getRootNode() { return fakeDocument; }
  get ownerDocument() { return fakeDocument; }
}
const fakeDocument = { nodeType: 9, adoptedStyleSheets: [], defaultView: { CSSStyleSheet: class { replaceSync() {} } }, createElement: tag => new FakeElement(tag) };

const REAL_SET_TIMEOUT = globalThis.setTimeout, REAL_CLEAR_TIMEOUT = globalThis.clearTimeout;
function installDom(t) {
  const previous = globalThis.document;
  globalThis.document = fakeDocument;
  t.after(() => { globalThis.document = previous; globalThis.setTimeout = REAL_SET_TIMEOUT; globalThis.clearTimeout = REAL_CLEAR_TIMEOUT; });
}
/** A fake clock (restored after the test by installDom). */
function installClock() {
  const timers = new Map(); let now = 0, next = 1;
  globalThis.setTimeout = (callback, delay = 0) => { const handle = next++; timers.set(handle, { at: now + delay, callback }); return handle; };
  globalThis.clearTimeout = handle => { timers.delete(handle); };
  return { get live() { return timers.size; }, tick(ms) { now += ms; for (const [handle, timer] of [...timers].sort((a, b) => a[1].at - b[1].at)) if (timer.at <= now && timers.has(handle)) { timers.delete(handle); timer.callback(); } } };
}

class FakeWorker {
  listeners = { message: new Set(), error: new Set(), messageerror: new Set() };
  posted = []; terminated = false;
  addEventListener(type, listener) { this.listeners[type].add(listener); }
  removeEventListener(type, listener) { this.listeners[type].delete(listener); }
  postMessage(data) { this.posted.push(data); }
  terminate() { this.terminated = true; }
  get listenerCount() { return Object.values(this.listeners).reduce((sum, set) => sum + set.size, 0); }
  emit(type, data) { for (const listener of [...this.listeners[type]]) listener(type === "message" ? { data } : { type }); }
  ready() { this.emit("message", { v: 1, type: "ready" }); }
  answer(model) { this.emit("message", { v: 1, type: "result", id: this.posted.at(-1).id, model }); }
}

/** The built-in DOCX renderer (so worker mode applies), recording which path rendered what. */
class RecordingDocx extends DocxRenderer {
  calls = [];
  render(context) { this.calls.push(["render", context.file.name]); return { element: new FakeElement("main-thread"), destroy() {} }; }
  renderParsed(model, context) { this.calls.push(["renderParsed", model.title, context.file.name]); return { element: new FakeElement("worker"), destroy() {} }; }
}
const docxBytes = new Uint8Array(buildFixture(ooxmlFixtures.find(fixture => fixture.id === "docx-basic")));
const docx = (name = "a.docx") => ({ name, source: docxBytes });
const model = title => ({ kind: "word", title, blocks: [{ type: "paragraph", value: { runs: [{ text: title }] } }] });
const settle = () => new Promise(resolve => setImmediate(resolve));

/** A viewer with the recording renderer and, unless `worker` is false, worker mode with fake workers. */
function viewer(t, { worker = {}, renderer = new RecordingDocx(), registry } = {}) {
  installDom(t);
  const workers = [], container = new FakeElement("div");
  const reg = registry ?? new RendererRegistry();
  if (!registry) reg.register(renderer);
  const options = { container, file: docx(), toolbar: false, registry: reg };
  if (worker !== false) options.worker = { create: () => { const created = new FakeWorker(); workers.push(created); return created; }, ...worker };
  const previewer = new FilePreviewer(options);
  t.after(() => previewer.destroy());
  const root = container.children.find(child => child.className === "ufp-root");
  return { previewer, workers, renderer, status: root.children[1], content: root.children[2] };
}

test("worker mode parses in the worker and renders the model with renderParsed; the input stays usable", async t => {
  const { previewer, workers, renderer, status, content } = viewer(t);
  const loading = previewer.open();
  await settle();
  assert.equal(workers.length, 1);
  workers[0].ready(); await settle();
  const request = workers[0].posted[0];
  assert.equal(request.format, "docx"); assert.equal(request.name, "a.docx");
  assert.equal(request.bytes.byteLength, docxBytes.byteLength);
  workers[0].answer(model("from the worker"));
  await loading;
  assert.deepEqual(renderer.calls, [["renderParsed", "from the worker", "a.docx"]]);
  assert.equal(content.children[0].tagName, "worker");
  assert.equal(status.textContent, "a.docx - DOCX");
  assert.equal(docxBytes.byteLength > 0, true, "the caller's bytes were copied, not transferred");
});

test("without the worker option, and for renderers the application registers, parsing stays on the main thread", async t => {
  const plain = viewer(t, { worker: false });
  await plain.previewer.open();
  assert.deepEqual(plain.renderer.calls, [["render", "a.docx"]]);
  // An application's renderer, even one with a renderParsed method, is never given a worker model.
  const registry = new RendererRegistry(), calls = [];
  registry.register({ formats: ["docx"], render: context => { calls.push("render"); return { element: new FakeElement("custom") }; }, renderParsed: () => { calls.push("renderParsed"); return { element: new FakeElement("x") }; } });
  const custom = viewer(t, { registry });
  await custom.previewer.open();
  assert.deepEqual(calls, ["render"]);
  assert.equal(custom.workers.length, 0, "no worker is created for files it does not parse");
});

test("a start-up failure fails the load by default and falls back to the main thread when allowed", async t => {
  const failing = viewer(t, { worker: { create: () => { throw new Error("blocked by CSP"); } } });
  await assert.rejects(failing.previewer.open(), error => error instanceof WorkerStartupError && error.code === "renderer-error");
  assert.equal(failing.status.className, "ufp-status ufp-status-error");
  assert.deepEqual(failing.renderer.calls, []);

  const fallback = viewer(t, { worker: { create: () => { throw new Error("blocked by CSP"); }, fallback: "main-thread" } });
  await fallback.previewer.open();
  assert.deepEqual(fallback.renderer.calls, [["render", "a.docx"]]);
  assert.equal(fallback.status.textContent, "a.docx - DOCX");
});

test("the start-up timeout falls back when allowed; the parse timeout never does", async t => {
  installDom(t);
  const clock = installClock();
  const startup = viewer(t, { worker: { fallback: "main-thread", startupTimeout: 1000 } });
  const loading = startup.previewer.open(); await settle();
  clock.tick(1000);
  await loading;
  assert.deepEqual(startup.renderer.calls, [["render", "a.docx"]]);
  assert.equal(startup.workers[0].terminated, true);

  const parse = viewer(t, { worker: { fallback: "main-thread", timeout: 2000 } });
  const parsing = parse.previewer.open(); await settle();
  parse.workers[0].ready(); await settle();
  clock.tick(2000);
  await assert.rejects(parsing, error => error.code === "resource-limit" && /did not finish within 2,000 ms/.test(error.message));
  assert.deepEqual(parse.renderer.calls, [], "no main-thread parse after a parse timeout");
  assert.equal(parse.status.textContent, "Parsing did not finish within 2,000 ms.");
  assert.equal(parse.workers[0].terminated, true);
});

test("cancelling while the worker starts never falls back to the main thread", async t => {
  const { previewer, workers, renderer, status } = viewer(t, { worker: { fallback: "main-thread" } });
  const controller = new AbortController();
  const loading = previewer.load(docx(), { signal: controller.signal });
  await settle();
  controller.abort();
  await loading; // cancelled loads resolve without showing an error
  assert.deepEqual(renderer.calls, []);
  assert.equal(workers[0].terminated, true);
  assert.equal(status.className, "ufp-status");
  assert.equal(status.textContent, "Preview loading was cancelled.");
});

test("a signal, a new load and destroy() each cancel a parse in flight by ending its worker", async t => {
  const { previewer, workers, renderer, content } = viewer(t);
  const controller = new AbortController();
  const first = previewer.load(docx("first.docx"), { signal: controller.signal });
  await settle(); workers[0].ready(); await settle();
  controller.abort();
  await first;
  assert.equal(workers[0].terminated, true);
  assert.deepEqual(renderer.calls, []);

  // A new load replaces the one in flight; its worker is ended and a new one parses the new file.
  const second = previewer.load(docx("second.docx"));
  await settle(); workers[1].ready(); await settle();
  const third = previewer.load(docx("third.docx"));
  await second;
  assert.equal(workers[1].terminated, true);
  await settle(); workers[2].ready(); await settle();
  assert.equal(workers[2].posted[0].name, "third.docx");
  workers[2].answer(model("third"));
  await third;
  assert.deepEqual(renderer.calls, [["renderParsed", "third", "third.docx"]]);
  assert.equal(content.children[0].tagName, "worker");

  // destroy() during a parse (the idle worker is reused for the fourth file).
  const fourth = previewer.load(docx("fourth.docx"));
  await settle();
  assert.equal(workers.length, 3); assert.equal(workers[2].posted.at(-1).name, "fourth.docx");
  previewer.destroy();
  await fourth;
  for (const worker of workers) { assert.equal(worker.terminated, true); assert.equal(worker.listenerCount, 0); }
  // The caller's signal no longer has a listener after its load ended.
  controller.abort();
});

test("a result that arrives as a newer load starts is never rendered (load-generation guard)", async t => {
  const { previewer, workers, renderer } = viewer(t);
  const older = previewer.load(docx("older.docx"));
  await settle(); workers[0].ready(); await settle();
  workers[0].answer(model("older"));
  // The older load resumes only after this synchronous call: by then it is no longer current.
  const newer = previewer.load(docx("newer.docx"));
  await older;
  assert.deepEqual(renderer.calls, [], "the older model was not rendered");
  await settle();
  const worker = workers.at(-1);
  if (worker.posted.length === 0) { worker.ready(); await settle(); }
  worker.answer(model("newer"));
  await newer;
  assert.deepEqual(renderer.calls, [["renderParsed", "newer", "newer.docx"]]);
});

test("worker errors and invalid models are shown as errors; the message reaches the page only as text", async t => {
  const { previewer, workers, status } = viewer(t);
  const loading = previewer.open();
  await settle(); workers[0].ready(); await settle();
  workers[0].emit("message", { v: 1, type: "error", id: 1, error: { code: "invalid-file", message: "<img src=x onerror=alert(1)>" } });
  await assert.rejects(loading, error => error instanceof PreviewError && error.code === "invalid-file");
  assert.equal(status.textContent, "<img src=x onerror=alert(1)>");
  assert.equal(status.children.length, 0);
  const second = previewer.open();
  await settle();
  workers.at(-1).answer({ kind: "presentation", title: "wrong", slides: [] });
  await assert.rejects(second, error => error.code === "renderer-error");
});

test("worker options and the load signal are checked", async t => {
  installDom(t);
  const base = { container: new FakeElement("div"), file: docx(), toolbar: false };
  assert.throws(() => new FilePreviewer({ ...base, worker: "worker.js" }), TypeError);
  assert.throws(() => new FilePreviewer({ ...base, worker: {} }), TypeError);
  assert.throws(() => new FilePreviewer({ ...base, worker: { create: () => new FakeWorker(), fallback: "silent" } }), RangeError);
  assert.throws(() => new FilePreviewer({ ...base, worker: { create: () => new FakeWorker(), timeout: 500 } }), RangeError);
  assert.throws(() => new FilePreviewer({ ...base, worker: { create: () => new FakeWorker(), startupTimeout: 120_000 } }), RangeError);
  const previewer = new FilePreviewer({ ...base, worker: { create: () => new FakeWorker() } });
  await assert.rejects(previewer.load(docx(), { signal: "stop" }), TypeError);
  previewer.destroy();
});
