import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { WorkerClient, WorkerStartupError } from "../dist/worker/worker-client.js";
import { errorMessage, MAX_ERROR_MESSAGE_LENGTH, readParseRequest, readWorkerMessage, validateModel } from "../dist/worker/protocol.js";
import { PreviewError } from "../dist/errors.js";
import { parseDocx, parsePptx, parseXlsx } from "../dist/parsers/ooxml.js";
import { parseOds } from "../dist/parsers/odf/ods.js";
import { writeDemoSamples } from "../scripts/demo-samples.mjs";
import { buildFixture, ooxmlFixtures } from "./helpers/ooxml-fixtures.mjs";
import { makeOds } from "./helpers/odf.mjs";

// The worker client (section 3 of the Phase 2 design) against an injected fake worker and a fake clock:
// start-up and parse timeouts, error and messageerror before and after "ready", cancellation in each state,
// destroy(), stale and invalid ids, overlapping parses, and no timers or listeners left behind.

const REAL_SET_TIMEOUT = globalThis.setTimeout, REAL_CLEAR_TIMEOUT = globalThis.clearTimeout;
/** A fake clock replacing setTimeout and clearTimeout, so tests can advance time and count live timers. The real functions are restored after the test. */
function installClock(t) {
  const timers = new Map();
  let now = 0, nextHandle = 1;
  globalThis.setTimeout = (callback, delay = 0) => { const handle = nextHandle++; timers.set(handle, { at: now + delay, callback }); return handle; };
  globalThis.clearTimeout = handle => { timers.delete(handle); };
  t.after(() => { globalThis.setTimeout = REAL_SET_TIMEOUT; globalThis.clearTimeout = REAL_CLEAR_TIMEOUT; });
  return {
    get live() { return timers.size; },
    tick(ms) {
      now += ms;
      for (const [handle, timer] of [...timers].sort((a, b) => a[1].at - b[1].at)) if (timer.at <= now && timers.has(handle)) { timers.delete(handle); timer.callback(); }
    },
  };
}

class FakeWorker {
  listeners = { message: new Set(), error: new Set(), messageerror: new Set() };
  posted = [];
  terminated = false;
  addEventListener(type, listener) { this.listeners[type].add(listener); }
  removeEventListener(type, listener) { this.listeners[type].delete(listener); }
  postMessage(data, transfer) { this.posted.push({ data, transfer }); }
  terminate() { this.terminated = true; }
  get listenerCount() { return Object.values(this.listeners).reduce((sum, set) => sum + set.size, 0); }
  emit(type, data) { for (const listener of [...this.listeners[type]]) listener(type === "message" ? { data } : { type }); }
  ready() { this.emit("message", { v: 1, type: "ready" }); }
  result(id, model = wordModel()) { this.emit("message", { v: 1, type: "result", id, model }); }
}

/** A client whose create() hands out FakeWorkers, recorded in `workers`. */
function setup(t, options = {}) {
  const clock = installClock(t), workers = [];
  const client = new WorkerClient({ create: () => { const worker = new FakeWorker(); workers.push(worker); return worker; }, maxFileSize: 1024, ...options });
  return { clock, workers, client, get worker() { return workers.at(-1); } };
}
/** Lets pending promise continuations run (the fake clock does not affect setImmediate). */
const settle = () => new Promise(resolve => setImmediate(resolve));
const wordModel = () => ({ kind: "word", title: "t.docx", blocks: [{ type: "paragraph", value: { runs: [{ text: "hello" }] } }] });
const bytes = () => new Uint8Array([1, 2, 3]);
const code = (expected, pattern = /./) => error => error instanceof PreviewError && error.code === expected && pattern.test(error.message);
const startup = (expected, pattern = /./) => error => error instanceof WorkerStartupError && error.code === expected && pattern.test(error.message);
/** A fake AbortSignal that counts its listeners. */
function fakeSignal() {
  const listeners = new Set();
  return { aborted: false, reason: undefined, addEventListener(type, listener) { listeners.add(listener); }, removeEventListener(type, listener) { listeners.delete(listener); }, abort() { this.aborted = true; for (const listener of [...listeners]) listener(); }, get listenerCount() { return listeners.size; } };
}

test("a parse waits for ready, transfers a copy of exactly the input bytes and returns the validated model", async t => {
  const { client, workers, clock } = setup(t);
  const backing = new Uint8Array([9, 9, 1, 2, 3, 9]), input = backing.subarray(2, 5);
  const result = client.parse("docx", input, "t.docx", { locale: "de-DE" });
  assert.equal(client.state, "starting");
  assert.equal(workers[0].posted.length, 0, "nothing is posted before ready");
  workers[0].ready(); await settle();
  assert.equal(client.state, "busy");
  const { data, transfer } = workers[0].posted[0];
  assert.deepEqual(Object.keys(data).sort(), ["bytes", "format", "id", "locale", "name", "type", "v"]);
  assert.deepEqual({ ...data, bytes: undefined }, { v: 1, type: "parse", id: 1, format: "docx", name: "t.docx", locale: "de-DE", bytes: undefined });
  assert.ok(data.bytes instanceof ArrayBuffer && transfer.length === 1 && transfer[0] === data.bytes);
  assert.deepEqual([...new Uint8Array(data.bytes)], [1, 2, 3], "only the view's bytes, in a buffer of their own");
  assert.notEqual(data.bytes, backing.buffer);
  assert.deepEqual([...backing], [9, 9, 1, 2, 3, 9], "the caller's array is untouched");
  workers[0].result(1);
  assert.deepEqual(await result, wordModel());
  assert.equal(client.state, "idle");
  assert.equal(clock.live, 0, "no timer left after the parse");
  // The worker is reused; ids increase.
  const second = client.parse("docx", bytes(), "t.docx");
  await settle();
  assert.equal(workers.length, 1);
  assert.equal(workers[0].posted[1].data.id, 2);
  workers[0].result(2);
  await second;
  client.destroy();
  assert.equal(workers[0].terminated, true);
  assert.equal(workers[0].listenerCount, 0);
});

test("error responses keep their known code; unknown codes become renderer-error; messages are truncated", async t => {
  const { client, workers } = setup(t);
  const first = client.parse("xlsx", bytes(), "a.xlsx"); workers[0].ready(); await settle();
  workers[0].emit("message", { v: 1, type: "error", id: 1, error: { code: "resource-limit", message: "too big" } });
  await assert.rejects(first, code("resource-limit", /^too big$/));
  assert.equal(client.state, "idle", "an error response does not end the worker");
  const second = client.parse("xlsx", bytes(), "a.xlsx"); await settle();
  workers[0].emit("message", { v: 1, type: "error", id: 2, error: { code: "evil-code", message: "x".repeat(10_000) } });
  await assert.rejects(second, error => error.code === "renderer-error" && error.message.length === MAX_ERROR_MESSAGE_LENGTH);
  assert.equal(workers.length, 1);
  client.destroy();
});

test("create() throwing or returning something other than a worker is a start-up failure", async t => {
  installClock(t);
  const throwing = new WorkerClient({ create: () => { throw new Error("CSP blocked the worker"); }, maxFileSize: 10 });
  await assert.rejects(throwing.parse("docx", bytes(), "a.docx"), startup("renderer-error", /could not be created/));
  assert.equal(throwing.state, "none");
  const wrong = new WorkerClient({ create: () => ({}), maxFileSize: 10 });
  await assert.rejects(wrong.parse("docx", bytes(), "a.docx"), startup("renderer-error"));
});

test("the start-up timeout is a start-up failure; the parse timeout is not, and each runs independently", async t => {
  const { client, workers, clock } = setup(t, { startupTimeout: 2000, timeout: 5000 });
  const starting = client.parse("docx", bytes(), "a.docx");
  clock.tick(1999); assert.equal(client.state, "starting");
  clock.tick(1);
  await assert.rejects(starting, startup("resource-limit", /did not start within 2,000 ms/));
  assert.equal(workers[0].terminated, true); assert.equal(workers[0].listenerCount, 0); assert.equal(clock.live, 0);

  const parsing = client.parse("docx", bytes(), "a.docx");
  assert.equal(workers.length, 2, "a new worker after the failure");
  clock.tick(1500); workers[1].ready(); await settle();
  clock.tick(4999); assert.equal(client.state, "busy", "the parse timeout starts when the request is posted");
  clock.tick(1);
  await assert.rejects(parsing, error => code("resource-limit", /did not finish within 5,000 ms/)(error) && !(error instanceof WorkerStartupError));
  assert.equal(workers[1].terminated, true); assert.equal(clock.live, 0);
  assert.equal(client.state, "none");
});

test("error and messageerror are start-up failures before ready and parse failures afterwards", async t => {
  for (const type of ["error", "messageerror"]) {
    const { client, workers, clock } = setup(t);
    const starting = client.parse("pptx", bytes(), "a.pptx");
    workers[0].emit(type);
    await assert.rejects(starting, startup("renderer-error"), `${type} before ready`);
    assert.equal(workers[0].terminated, true); assert.equal(workers[0].listenerCount, 0);
    const parsing = client.parse("pptx", bytes(), "a.pptx");
    workers[1].ready(); await settle();
    workers[1].emit(type);
    await assert.rejects(parsing, error => code("renderer-error")(error) && !(error instanceof WorkerStartupError), `${type} after ready`);
    assert.equal(workers[1].terminated, true); assert.equal(clock.live, 0);
    // In the idle state the worker is simply ended; the next parse creates a new one.
    const idle = client.parse("pptx", bytes(), "a.pptx");
    workers[2].ready(); await settle(); workers[2].result(workers[2].posted[0].data.id, { kind: "presentation", title: "t", slides: [] }); await idle;
    workers[2].emit(type);
    assert.equal(workers[2].terminated, true); assert.equal(client.state, "none");
    client.destroy();
  }
});

test("invalid messages before ready are start-up failures; afterwards they end the parse with renderer-error", async t => {
  const invalidBeforeReady = [{ v: 1, type: "result", id: 1, model: {} }, { v: 2, type: "ready" }, { v: 1, type: "ready", extra: 1 }, "ready", null];
  for (const message of invalidBeforeReady) {
    const { client, workers } = setup(t);
    const starting = client.parse("docx", bytes(), "a.docx");
    workers[0].emit("message", message);
    await assert.rejects(starting, startup("renderer-error", /invalid message/), JSON.stringify(message));
    assert.equal(workers[0].terminated, true);
  }
  const invalidAfterReady = [
    { v: 1, type: "ready" },
    { v: 1, type: "result", id: 7, model: wordModel() },
    { v: 1, type: "result", id: 0, model: wordModel() },
    { v: 1, type: "result", id: 1.5, model: wordModel() },
    { v: 1, type: "result", id: 1, model: wordModel(), extra: true },
    { v: 1, type: "error", id: 1, error: { code: "aborted" } },
    { v: 1, type: "progress", id: 1 },
    { v: "1", type: "result", id: 1, model: wordModel() },
  ];
  for (const message of invalidAfterReady) {
    const { client, workers, clock } = setup(t);
    const parsing = client.parse("docx", bytes(), "a.docx");
    workers[0].ready(); await settle();
    workers[0].emit("message", message);
    await assert.rejects(parsing, error => code("renderer-error", /invalid message/)(error) && !(error instanceof WorkerStartupError), JSON.stringify(message));
    assert.equal(workers[0].terminated, true); assert.equal(clock.live, 0);
  }
});

test("a response for an issued id that is no longer pending is stale and ignored; a never-issued id ends the worker", async t => {
  const { client, workers } = setup(t);
  const first = client.parse("docx", bytes(), "a.docx"); workers[0].ready(); await settle();
  workers[0].result(1); await first;
  workers[0].result(1); workers[0].emit("message", { v: 1, type: "error", id: 1, error: { code: "parse-error", message: "late" } });
  assert.equal(workers[0].terminated, false, "stale responses are ignored");
  assert.equal(client.state, "idle");
  const second = client.parse("docx", bytes(), "a.docx"); await settle();
  workers[0].result(1); // stale while id 2 is pending
  assert.equal(client.state, "busy");
  workers[0].result(2); await second;
  workers[0].result(3); // never issued
  assert.equal(workers[0].terminated, true);
  assert.equal(client.state, "none");
});

test("cancellation during start-up rejects with aborted (never a start-up failure) and removes its listener", async t => {
  const { client, workers, clock } = setup(t);
  const signal = fakeSignal();
  const starting = client.parse("docx", bytes(), "a.docx", { signal });
  assert.equal(signal.listenerCount, 1);
  signal.abort();
  await assert.rejects(starting, error => code("aborted")(error) && !(error instanceof WorkerStartupError));
  assert.equal(workers[0].terminated, true); assert.equal(workers[0].listenerCount, 0);
  assert.equal(signal.listenerCount, 0); assert.equal(clock.live, 0);
  // An already aborted signal: nothing is created.
  const already = fakeSignal(); already.aborted = true;
  await assert.rejects(client.parse("docx", bytes(), "a.docx", { signal: already }), code("aborted"));
  assert.equal(workers.length, 1);
});

test("cancellation during parsing terminates the worker; the next parse creates a new one", async t => {
  const { client, workers, clock } = setup(t);
  const signal = fakeSignal();
  const parsing = client.parse("docx", bytes(), "a.docx", { signal });
  workers[0].ready(); await settle();
  signal.abort();
  await assert.rejects(parsing, code("aborted"));
  assert.equal(workers[0].terminated, true); assert.equal(signal.listenerCount, 0); assert.equal(clock.live, 0);
  const next = client.parse("docx", bytes(), "a.docx");
  assert.equal(workers.length, 2);
  workers[1].ready(); await settle(); workers[1].result(workers[1].posted[0].data.id); await next;
  // A signal of a settled parse no longer affects the client.
  signal.abort();
  assert.equal(workers[1].terminated, false);
  client.destroy();
});

test("a new parse cancels the one in flight, also while it waits for ready or right after ready", async t => {
  const { client, workers } = setup(t);
  const first = client.parse("docx", bytes(), "a.docx");
  const second = client.parse("docx", bytes(), "b.docx");
  await assert.rejects(first, code("aborted"));
  assert.equal(workers.length, 2); assert.equal(workers[0].terminated, true);
  workers[1].ready(); await settle(); workers[1].result(1); await second;
  // In flight: replaced.
  const third = client.parse("docx", bytes(), "c.docx"); await settle();
  const fourth = client.parse("docx", bytes(), "d.docx");
  await assert.rejects(third, code("aborted"));
  workers[2].ready(); await settle();
  assert.equal(workers[2].posted[0].data.name, "d.docx");
  workers[2].result(workers[2].posted[0].data.id); await fourth;
  // A newer parse called after ready but before the older one resumed: the older one gives up, the newer one is posted.
  client.destroy();
  const { client: other, workers: others } = setup(t);
  const older = other.parse("docx", bytes(), "old.docx");
  others[0].ready();
  const newer = other.parse("docx", bytes(), "new.docx");
  await assert.rejects(older, code("aborted"));
  assert.deepEqual(others[0].posted.map(entry => entry.data.name), ["new.docx"]);
  others[0].result(others[0].posted[0].data.id); await newer;
  other.destroy();
});

test("destroy() in every state rejects with aborted, leaves no timer or listener, and is idempotent", async t => {
  for (const state of ["none", "starting", "idle", "busy"]) {
    const { client, workers, clock } = setup(t);
    let pending;
    if (state !== "none") pending = client.parse("docx", bytes(), "a.docx");
    if (state === "idle" || state === "busy") { workers[0].ready(); await settle(); }
    if (state === "idle") { workers[0].result(1); await pending; pending = undefined; }
    assert.equal(client.state, state);
    client.destroy(); client.destroy();
    if (pending) await assert.rejects(pending, code("aborted"), state);
    for (const worker of workers) { assert.equal(worker.terminated, true, state); assert.equal(worker.listenerCount, 0, state); }
    assert.equal(clock.live, 0, `${state}: timers left`);
    await assert.rejects(client.parse("docx", bytes(), "a.docx"), code("aborted"), `${state}: parse after destroy`);
    assert.equal(workers.length, state === "none" ? 0 : 1, "destroy() never creates a worker");
  }
});

test("models are validated before they are returned: output limits, plain data, kind and rendering paths", async t => {
  const cases = [
    ["a model of another kind", { kind: "workbook", title: "t", sheets: [] }, code("renderer-error", /kind/)],
    ["an unknown block type", { kind: "word", title: "t", blocks: [{ type: "script", value: {} }] }, code("renderer-error", /blocks\[0\]\.type/)],
    ["a string over the length limit", { kind: "word", title: "x".repeat(1024 * 1024 + 1), blocks: [] }, code("resource-limit", /string length/)],
    ["a class instance", { kind: "word", title: "t", blocks: [], when: new Date(0) }, code("renderer-error", /invalid document model/)],
    ["a non-finite number", { kind: "word", title: "t", blocks: [], page: { width: Infinity } }, code("renderer-error")],
  ];
  const cyclic = wordModel(); cyclic.blocks[0].value.self = cyclic.blocks[0];
  cases.push(["a cycle", cyclic, code("renderer-error")]);
  for (const [label, model, expected] of cases) {
    const { client, workers } = setup(t);
    const parsing = client.parse("docx", bytes(), "a.docx"); workers[0].ready(); await settle();
    workers[0].result(1, model);
    await assert.rejects(parsing, expected, label);
    assert.equal(client.state, "idle", `${label}: the worker answered correctly, so it is kept`);
    client.destroy();
  }
});

test("options are checked: create must be a function; timeouts are whole milliseconds within their ranges; inputs over maxFileSize are refused", async t => {
  installClock(t);
  const create = () => new FakeWorker();
  assert.throws(() => new WorkerClient({ create: "worker.js", maxFileSize: 1 }), TypeError);
  for (const startupTimeout of [999, 60_001, 1.5, NaN]) assert.throws(() => new WorkerClient({ create, maxFileSize: 1, startupTimeout }), RangeError, String(startupTimeout));
  for (const timeout of [999, 300_001, Infinity]) assert.throws(() => new WorkerClient({ create, maxFileSize: 1, timeout }), RangeError, String(timeout));
  assert.doesNotThrow(() => new WorkerClient({ create, maxFileSize: 1, startupTimeout: 1000, timeout: 300_000 }));
  let created = 0;
  const client = new WorkerClient({ create: () => { created++; return new FakeWorker(); }, maxFileSize: 2 });
  await assert.rejects(client.parse("docx", bytes(), "a.docx"), code("resource-limit"));
  assert.equal(created, 0);
});

test("the worker side reads only exact parse requests", () => {
  const buffer = new ArrayBuffer(4), base = { v: 1, type: "parse", id: 3, format: "ods", name: "a.ods", bytes: buffer };
  assert.deepEqual(readParseRequest(base), base);
  assert.deepEqual(readParseRequest({ ...base, locale: "en-US" }), { ...base, locale: "en-US" });
  const invalid = [
    { ...base, v: 2 }, { ...base, type: "result" }, { ...base, id: 0 }, { ...base, id: "3" }, { ...base, format: "svg" }, { ...base, format: "pdf" },
    { ...base, name: 5 }, { ...base, name: "n".repeat(1025) }, { ...base, locale: "en US" }, { ...base, locale: "x".repeat(65) }, { ...base, locale: "<script>" },
    { ...base, bytes: new Uint8Array(4) }, { ...base, bytes: "AAAA" }, { ...base, extra: 1 }, Object.assign(Object.create({ polluted: true }), base), null, [], "parse",
  ];
  for (const message of invalid) assert.equal(readParseRequest(message), undefined, JSON.stringify(message));
  const { bytes: omitted, ...withoutBytes } = base;
  assert.equal(readParseRequest(withoutBytes), undefined);
  assert.equal(readWorkerMessage({ v: 1, type: "ready" }).type, "ready");
});

test("errors leave the worker as a known code and a bounded message only", () => {
  assert.deepEqual(errorMessage(2, new PreviewError("bad zip", "invalid-file")), { v: 1, type: "error", id: 2, error: { code: "invalid-file", message: "bad zip" } });
  // Anything else (a TypeError from a bug, a value thrown) is not passed on in detail.
  assert.deepEqual(errorMessage(2, new TypeError("secret internals")), { v: 1, type: "error", id: 2, error: { code: "renderer-error", message: "The file could not be previewed." } });
  assert.equal(errorMessage(1, new PreviewError("y".repeat(900), "parse-error")).error.message.length, MAX_ERROR_MESSAGE_LENGTH);
});

test("every fixture and demo model passes the main-thread validation for its format", async () => {
  const parsers = { docx: parseDocx, pptx: parsePptx, xlsx: parseXlsx, ods: parseOds };
  const models = ooxmlFixtures.map(fixture => [fixture.id, fixture.kind, parsers[fixture.kind](buildFixture(fixture), fixture.file, fixture.options)]);
  const directory = mkdtempSync(join(tmpdir(), "ufp-protocol-"));
  try {
    await writeDemoSamples(directory);
    for (const [file, kind] of [["sample-report.docx", "docx"], ["sample-workbook.xlsx", "xlsx"], ["sample-deck.pptx", "pptx"]]) models.push([file, kind, parsers[kind](new Uint8Array(readFileSync(join(directory, file))), file)]);
  } finally { rmSync(directory, { recursive: true, force: true }); }
  models.push(["ods", "ods", parseOds(makeOds({ body: '<table:table table:name="S"><table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row></table:table>' }), "s.ods")]);
  for (const [id, format, model] of models) {
    // The model as the worker sends it: a structured clone.
    assert.doesNotThrow(() => validateModel(structuredClone(model), format), id);
  }
  assert.throws(() => validateModel(models.find(([, format]) => format === "docx")[2], "xlsx"), code("renderer-error", /kind/));
});
