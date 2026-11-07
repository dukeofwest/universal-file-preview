import test from "node:test";
import assert from "node:assert/strict";
import { Worker } from "node:worker_threads";
import { dosCorpus } from "./helpers/dos-corpus.mjs";

// Security audit acceptance criterion 8: every file of the denial-of-service corpus is refused with a
// PreviewError or handled within the library's limits, in under 2 seconds and with a JavaScript heap capped at
// 256 MiB, unless its corpus entry states a larger budget and the reason. Each file runs alone in a worker thread
// whose old generation is capped: a file that needs more ends that thread with an out-of-memory error, which fails
// the test. UFP_DOS_TIME_LIMIT_MS and UFP_DOS_HEAP_LIMIT_MIB change the defaults for measurements.
// Timing-sensitive, so it is not part of the parallel `npm test` run: `npm run test:security` runs it alone (CI too).

const TIME_LIMIT_MS = Number(process.env.UFP_DOS_TIME_LIMIT_MS) || 2000;
const HEAP_LIMIT_MIB = Number(process.env.UFP_DOS_HEAP_LIMIT_MIB) || 256;

/** The time budget of an entry: the default, or the larger budget the entry states. */
const timeLimit = entry => Math.max(TIME_LIMIT_MS, entry.budget?.ms ?? 0);

function runIsolated(entry, bytes) {
  return new Promise(resolve => {
    const heap = Math.max(HEAP_LIMIT_MIB, entry.budget?.heapMiB ?? 0);
    const worker = new Worker(new URL("./helpers/dos-worker.mjs", import.meta.url), { resourceLimits: { maxOldGenerationSizeMb: heap } });
    let settled = false;
    // Watchdog: a file that never finishes (an endless loop) would otherwise stop this test, and CI, for good.
    const watchdogMs = 2 * timeLimit(entry) + 5000;
    const watchdog = setTimeout(() => finish({ outcome: "hang", message: `no result within ${watchdogMs} ms; the worker was terminated` }), watchdogMs);
    const finish = result => { if (settled) return; settled = true; clearTimeout(watchdog); void worker.terminate(); resolve(result); };
    worker.once("message", finish);
    worker.once("error", error => finish({ outcome: "crashed", code: error.code, message: error.message }));
    worker.once("exit", code => finish({ outcome: "crashed", message: `worker exited with code ${code}` }));
    worker.postMessage({ kind: entry.kind, name: entry.id, bytes }, [bytes.buffer]);
  });
}

test("the watchdog ends a worker that never answers and reports a hang", async () => {
  const started = performance.now();
  const result = await runIsolated({ id: "watchdog self-test", kind: "watchdog-self-test" }, new Uint8Array(1));
  assert.equal(result.outcome, "hang", JSON.stringify(result));
  assert.ok(performance.now() - started < 2 * TIME_LIMIT_MS + 15_000, "the watchdog fired");
});

test("every denial-of-service corpus file is refused or handled in time, within a capped heap", async t => {
  const results = [];
  for (const entry of dosCorpus()) {
    const bytes = entry.build(), size = bytes.byteLength;
    results.push({ entry, size, result: await runIsolated(entry, bytes) });
  }
  const lines = results.map(({ entry, result, size }) => {
    const budget = entry.budget ? ` [budget ${entry.budget.ms} ms, ${entry.budget.heapMiB} MiB: ${entry.budget.reason}]` : "";
    const outcome = `${result.outcome}${result.code ? ` (${result.code})` : ""}`;
    return `${entry.id}${budget}: ${outcome} in ${result.ms ?? "?"} ms, heap ${result.heapUsedMiB ?? "?"} MiB, file ${(size / 1024).toFixed(0)} KiB${result.outcome === "crashed" || result.outcome === "hang" ? ` [${result.message}]` : ""}`;
  });
  t.diagnostic([`default budget ${TIME_LIMIT_MS} ms, ${HEAP_LIMIT_MIB} MiB`, ...lines].join(" | "));
  for (const { entry, result } of results) {
    assert.ok(result.outcome !== "crashed" && result.outcome !== "hang", `${entry.id}: ${result.outcome}: ${result.message}`);
    assert.equal(result.outcome, entry.expect, `${entry.id} (${entry.limit}): expected ${entry.expect}, got ${result.outcome} ${result.code ?? ""} ${result.message ?? ""}`);
    if (entry.expect === "refused") assert.ok(["resource-limit", "invalid-file", "parse-error"].includes(result.code), `${entry.id}: refused with ${result.code}`);
    const limit = timeLimit(entry);
    assert.ok(result.ms < limit, `${entry.id}: took ${result.ms} ms (limit ${limit} ms${entry.budget ? `: ${entry.budget.reason}` : ""})`);
  }
});
